import os
import io
import json
import hmac
import time
import base64
import hashlib
import tempfile
import traceback
from pathlib import Path
from dataclasses import asdict
from typing import Optional, List
from fastapi import FastAPI, UploadFile, File, Form, Query, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

from src.domain.models import HeadcountRecord
from src.domain.interfaces import IHeadcountRepository
from src.infrastructure.databricks_repository import (
    DatabricksHeadcountRepository,
    ExcelHeadcountRepository,
    parse_reference_date_info,
)
from src.services.analytics_service import PeopleAnalyticsService
from src.services.updater_service import PresentationUpdaterService

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_FILE_PATH = PROJECT_ROOT / ".env"
load_dotenv(dotenv_path=ENV_FILE_PATH, override=True)

MAPPING_JSON_PATH = PROJECT_ROOT / "config" / "slide_mapping.json"
WEB_INDEX_PATH = PROJECT_ROOT / "web" / "index.html"
TMP_UPLOAD_CACHE_PATH = Path(tempfile.gettempdir()) / "si_pa_uploaded_records.json"
TMP_AUDIT_LOG_PATH = Path(tempfile.gettempdir()) / "si_pa_session_audit_log.json"

COOKIE_NAME = "si_pa_session"
APP_SECRET = (os.getenv("APP_SECRET_KEY") or "cni-secret-key-2026").strip()

app = FastAPI(
    title="Sistema Indústria | People Analytics - Estrutura Organizacional",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

_HEADCOUNT_CACHE: dict[str, tuple[List[HeadcountRecord], str]] = {}
_CUSTOM_UPLOADED_DATA: Optional[tuple[List[HeadcountRecord], str]] = None


def _append_server_audit_log(username: str, action: str, details: str) -> None:
    try:
        log_list = []
        if TMP_AUDIT_LOG_PATH.exists():
            log_list = json.loads(TMP_AUDIT_LOG_PATH.read_text(encoding="utf-8"))

        timestamp = time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())
        log_list.insert(
            0,
            {
                "timestamp": timestamp,
                "username": username,
                "action": action,
                "details": details,
            },
        )
        if len(log_list) > 100:
            log_list = log_list[:100]
        TMP_AUDIT_LOG_PATH.write_text(json.dumps(log_list, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def _resolve_metadata_from_records(
    records: List[HeadcountRecord],
    fallback_ref: str = "",
    fallback_filename: str = "",
) -> dict[str, str]:
    if records:
        for r in records:
            if r.ano_mes_referencia and r.ano_mes_referencia != "--/----":
                info = parse_reference_date_info(r.ano_mes_referencia, fallback_filename=fallback_filename)
                if info["ano_mes"] != "--/----":
                    return info
    return parse_reference_date_info(fallback_ref, fallback_filename=fallback_filename)


def _save_uploaded_to_tmp(records: List[HeadcountRecord], source_label: str) -> None:
    try:
        payload = {
            "source_label": source_label,
            "records": [asdict(r) for r in records],
        }
        TMP_UPLOAD_CACHE_PATH.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def _load_uploaded_from_tmp() -> Optional[tuple[List[HeadcountRecord], str]]:
    try:
        if TMP_UPLOAD_CACHE_PATH.exists():
            raw = json.loads(TMP_UPLOAD_CACHE_PATH.read_text(encoding="utf-8"))
            recs = [HeadcountRecord(**item) for item in raw.get("records", [])]
            if recs:
                return recs, raw.get("source_label", "Upload Web (Cache)")
    except Exception:
        pass
    return None


def load_users_registry() -> dict[str, dict[str, str]]:
    registry: dict[str, dict[str, str]] = {}
    legacy_user = (os.getenv("APP_ADMIN_USER") or "admin.sdh").strip()
    legacy_pass = (os.getenv("APP_ADMIN_PASSWORD") or "SistemaIndustria@2026").strip()
    if legacy_user and legacy_pass:
        registry[legacy_user.lower()] = {
            "username": legacy_user,
            "password": legacy_pass,
            "role": "ADMIN",
        }

    raw_users = (os.getenv("APP_USERS") or "").strip()
    if raw_users:
        for entry in raw_users.split(";"):
            parts = [p.strip() for p in entry.split(":")]
            if len(parts) >= 2 and parts[0] and parts[1]:
                u_name = parts[0]
                u_pass = parts[1]
                u_role = parts[2].upper() if len(parts) >= 3 else "EDITOR"
                registry[u_name.lower()] = {
                    "username": u_name,
                    "password": u_pass,
                    "role": u_role,
                }
    return registry


def create_session_token(username: str, role: str) -> str:
    expires = int(time.time()) + (12 * 3600)
    payload = f"{username}:{role}:{expires}"
    sig = hmac.new(APP_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return f"{payload}:{sig}"


def verify_session_token(token: Optional[str]) -> Optional[dict[str, str]]:
    if not token or token.count(":") != 3:
        return None
    username, role, exp_str, sig = token.split(":")
    try:
        if int(time.time()) > int(exp_str):
            return None
    except ValueError:
        return None
    payload = f"{username}:{role}:{exp_str}"
    expected_sig = hmac.new(APP_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    if hmac.compare_digest(sig, expected_sig):
        return {"username": username, "role": role}
    return None


def require_auth(request: Request, allow_viewer: bool = True) -> dict[str, str]:
    user_info = verify_session_token(request.cookies.get(COOKIE_NAME))
    if not user_info:
        raise HTTPException(status_code=401, detail="Sessão expirada ou não autenticada.")
    if not allow_viewer and user_info["role"] == "VIEWER":
        raise HTTPException(
            status_code=403,
            detail="O seu perfil (VIEWER) possui permissão apenas para consulta de Headcount.",
        )
    return user_info


def get_repository() -> tuple[Optional[IHeadcountRepository], str]:
    db_repo = DatabricksHeadcountRepository()
    if db_repo.is_configured():
        return db_repo, "Databricks Unity Catalog (Live)"

    env_excel = (os.getenv("FALLBACK_EXCEL_PATH") or "").strip()
    if env_excel:
        excel_path = Path(env_excel)
        if not excel_path.is_absolute():
            excel_path = PROJECT_ROOT / excel_path
        if excel_path.exists():
            return ExcelHeadcountRepository(str(excel_path), excel_path.name), f"Base Local ({excel_path.name})"

    return None, "Aguardando Upload de Planilha (.xlsx/.csv)"


def fetch_cached_records(reference_date: str = "", force_refresh: bool = False) -> tuple[List[HeadcountRecord], str]:
    global _CUSTOM_UPLOADED_DATA

    repo, source_label = get_repository()
    if repo is None:
        if _CUSTOM_UPLOADED_DATA is not None:
            return _CUSTOM_UPLOADED_DATA
        tmp_loaded = _load_uploaded_from_tmp()
        if tmp_loaded is not None:
            _CUSTOM_UPLOADED_DATA = tmp_loaded
            return tmp_loaded
        return [], "Aguardando Upload de Planilha (.xlsx/.csv)"

    if not force_refresh:
        if _CUSTOM_UPLOADED_DATA is not None:
            return _CUSTOM_UPLOADED_DATA
        tmp_loaded = _load_uploaded_from_tmp()
        if tmp_loaded is not None:
            _CUSTOM_UPLOADED_DATA = tmp_loaded
            return tmp_loaded

    norm_key = parse_reference_date_info(reference_date)["iso_date"] or "LATEST"

    if force_refresh or norm_key not in _HEADCOUNT_CACHE:
        try:
            records = repo.fetch_current_headcount(reference_date)
            _HEADCOUNT_CACHE[norm_key] = (records, source_label)
            if force_refresh:
                _CUSTOM_UPLOADED_DATA = None
        except Exception:
            if _CUSTOM_UPLOADED_DATA is not None:
                return _CUSTOM_UPLOADED_DATA
            tmp_loaded = _load_uploaded_from_tmp()
            if tmp_loaded is not None:
                return tmp_loaded
            return [], "Aguardando Upload de Planilha (.xlsx/.csv)"

    return _HEADCOUNT_CACHE.get(norm_key, ([], "Aguardando Upload de Planilha (.xlsx/.csv)"))


@app.post("/api/login")
async def login(response: Response, username: str = Form(...), password: str = Form(...)):
    users_db = load_users_registry()
    user_record = users_db.get(username.strip().lower())

    if user_record and hmac.compare_digest(password.strip(), user_record["password"]):
        token = create_session_token(user_record["username"], user_record["role"])
        response.set_cookie(
            key=COOKIE_NAME,
            value=token,
            httponly=True,
            max_age=12 * 3600,
            samesite="lax",
        )
        _append_server_audit_log(user_record["username"], "LOGIN", "Sessão iniciada com sucesso.")
        return {
            "status": "authenticated",
            "user": user_record["username"],
            "role": user_record["role"],
        }
    raise HTTPException(status_code=401, detail="Utilizador ou palavra-passe inválidos.")


@app.post("/api/logout")
async def logout(request: Request, response: Response):
    user_info = verify_session_token(request.cookies.get(COOKIE_NAME))
    if user_info:
        _append_server_audit_log(user_info["username"], "LOGOUT", "Sessão encerrada.")
    response.delete_cookie(COOKIE_NAME)
    return {"status": "logged_out"}


@app.get("/api/session")
async def check_session(request: Request):
    user_info = verify_session_token(request.cookies.get(COOKIE_NAME))
    if not user_info:
        return {"authenticated": False, "user": None, "role": None}
    return {
        "authenticated": True,
        "user": user_info["username"],
        "role": user_info["role"],
    }


@app.get("/", response_class=HTMLResponse)
async def serve_dashboard():
    if not WEB_INDEX_PATH.exists():
        return HTMLResponse("<h1>Arquivo web/index.html não encontrado.</h1>", status_code=404)
    return HTMLResponse(content=WEB_INDEX_PATH.read_text(encoding="utf-8"))


@app.post("/api/upload-headcount-data")
async def upload_headcount_data_endpoint(
    request: Request,
    data_file: UploadFile = File(...),
):
    global _CUSTOM_UPLOADED_DATA
    user = require_auth(request, allow_viewer=False)

    fname = (data_file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xls") or fname.endswith(".csv")):
        raise HTTPException(
            status_code=400,
            detail="Formato inválido. Envie um arquivo .xlsx ou .csv extraído do Databricks.",
        )

    try:
        raw_bytes = await data_file.read()
        repo = ExcelHeadcountRepository(raw_bytes, filename=data_file.filename or "upload.xlsx")
        records = repo.fetch_current_headcount()

        meta = _resolve_metadata_from_records(records, "", data_file.filename or "")
        source_label = f"Upload Web ({data_file.filename})"
        _CUSTOM_UPLOADED_DATA = (records, source_label)
        if meta["iso_date"]:
            _HEADCOUNT_CACHE[meta["iso_date"]] = (records, source_label)
        _save_uploaded_to_tmp(records, source_label)

        _append_server_audit_log(
            user["username"],
            "UPLOAD_BASE",
            f"Arquivo '{data_file.filename}' carregado com sucesso. Competência: {meta['ano_mes']} ({len(records)} registros).",
        )

        summary = PeopleAnalyticsService.build_organization_summary(records=records)
        return JSONResponse(
            content={
                "status": "success",
                "source": source_label,
                "reference_date": meta["ano_mes"],
                "iso_reference_date": meta["iso_date"],
                "default_label": meta["label"],
                "total_records": len(records),
                "total_headcount": summary["total_headcount"],
                "data": summary,
            }
        )
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=400, detail=f"Erro ao ler planilha enviada: {str(exc)}")


@app.get("/api/headcount")
async def get_organization_headcount(
    request: Request,
    reference_date: Optional[str] = Query(None),
    segmento: Optional[str] = Query(None),
    area: Optional[str] = Query(None),
    search: Optional[str] = Query(None),
    refresh: bool = Query(False),
):
    require_auth(request, allow_viewer=True)
    ref_date = (reference_date or "").strip()
    try:
        records, source_label = fetch_cached_records(ref_date, force_refresh=refresh)
        meta = _resolve_metadata_from_records(
            records,
            fallback_ref=ref_date if records else "",
            fallback_filename=source_label if records else "",
        )
        summary = PeopleAnalyticsService.build_organization_summary(
            records=records,
            segmento_filter=segmento,
            area_filter=area,
            search_term=search,
        )
        return JSONResponse(
            content={
                "status": "success",
                "source": source_label,
                "reference_date": meta["ano_mes"],
                "iso_reference_date": meta["iso_date"],
                "default_label": meta["label"],
                "data": summary,
            }
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro de consulta: {str(exc)}")


@app.post("/api/update-pptx")
async def update_powerpoint_endpoint(
    request: Request,
    file: UploadFile = File(...),
    data_file: Optional[UploadFile] = File(None),
    reference_date: str = Form(""),
    reference_label: str = Form(""),
    custom_overrides_json: str = Form("[]"),
    auto_layout: bool = Form(True),
):
    user = require_auth(request, allow_viewer=False)
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Envie um arquivo válido com extensão .pptx")

    try:
        raw_bytes = await file.read()
        pptx_stream = io.BytesIO(raw_bytes)
        custom_overrides = json.loads(custom_overrides_json) if custom_overrides_json else []

        if data_file and data_file.filename:
            data_bytes = await data_file.read()
            temp_repo = ExcelHeadcountRepository(data_bytes, filename=data_file.filename)
            records = temp_repo.fetch_current_headcount()
            meta = _resolve_metadata_from_records(records, "", data_file.filename)
        else:
            records, source_label = fetch_cached_records(reference_date)
            meta = _resolve_metadata_from_records(records, reference_date, source_label)

        if not records:
            raise HTTPException(
                status_code=400,
                detail="Nenhuma base de Headcount ativa. Suba primeiro a planilha (.xlsx/.csv) no botão 'Subir Base'.",
            )

        clean_req_label = reference_label.strip()
        if clean_req_label and clean_req_label not in ("Base Não Identificada", "Aguardando Base"):
            effective_label = clean_req_label
        else:
            effective_label = meta["label"]

        # O arquivo slide_mapping.json agora é 100% opcional: a leitura geométrica descobre tudo em memória
        updater = PresentationUpdaterService(mapping_json_path=str(MAPPING_JSON_PATH))
        updated_pptx_bytes, report = updater.execute_update(
            pptx_stream=pptx_stream,
            records=records,
            reference_label=effective_label,
            original_filename=file.filename,
            custom_overrides=custom_overrides,
            auto_layout=auto_layout,
        )

        _append_server_audit_log(
            user["username"],
            "GERAR_PPTX",
            f"Gerado arquivo '{report.output_filename}' ({len(custom_overrides)} ajustes aplicados, Total CNI: {report.total_company_headcount}).",
        )

        encoded_pptx = base64.b64encode(updated_pptx_bytes).decode("utf-8")
        return JSONResponse(
            content={
                "status": "success",
                "report": asdict(report),
                "pptx_base64": encoded_pptx,
                "download_filename": report.output_filename,
            }
        )
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Erro interno no processamento do PowerPoint: {str(exc)}")


@app.get("/api/audit-logs")
async def get_audit_logs_endpoint(request: Request):
    require_auth(request, allow_viewer=True)
    try:
        log_list = []
        if TMP_AUDIT_LOG_PATH.exists():
            log_list = json.loads(TMP_AUDIT_LOG_PATH.read_text(encoding="utf-8"))
        return JSONResponse(content={"status": "success", "logs": log_list})
    except Exception:
        return JSONResponse(content={"status": "success", "logs": []})


@app.get("/api/audit-logs/export")
async def export_audit_logs_txt(request: Request):
    require_auth(request, allow_viewer=True)
    try:
        log_list = []
        if TMP_AUDIT_LOG_PATH.exists():
            log_list = json.loads(TMP_AUDIT_LOG_PATH.read_text(encoding="utf-8"))

        lines = ["===================================================================="]
        lines.append(" SISTEMA INDÚSTRIA · LOG DE AUDITORIA E EXTRAÇÃO DE SESSÃO")
        lines.append("====================================================================\n")
        for item in log_list:
            lines.append(
                f"[{item.get('timestamp')}] ({item.get('username')}) {item.get('action')}: {item.get('details')}"
            )

        txt_content = "\n".join(lines)
        return PlainTextResponse(
            content=txt_content,
            media_type="text/plain;charset=utf-8",
            headers={"Content-Disposition": "attachment; filename=Log_Auditoria_Sistema_Industria.txt"},
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro ao exportar log: {str(exc)}")