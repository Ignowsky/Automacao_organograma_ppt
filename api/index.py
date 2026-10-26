import os
import io
import json
import hmac
import time
import base64
import hashlib
import traceback
from pathlib import Path
from dataclasses import asdict
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, Query, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

from src.domain.interfaces import IHeadcountRepository
from src.infrastructure.databricks_repository import (
    DatabricksHeadcountRepository,
    ExcelHeadcountRepository,
)
from src.services.analytics_service import PeopleAnalyticsService
from src.services.updater_service import PresentationUpdaterService

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_FILE_PATH = PROJECT_ROOT / ".env"
load_dotenv(dotenv_path=ENV_FILE_PATH, override=True)

DEFAULT_EXCEL_PATH = PROJECT_ROOT / "data" / "input" / "New_Query_2026_09_10_14_47_53 (1).xlsx"
MAPPING_JSON_PATH = PROJECT_ROOT / "config" / "slide_mapping.json"
WEB_INDEX_PATH = PROJECT_ROOT / "web" / "index.html"

COOKIE_NAME = "si_pa_session"
APP_SECRET = (os.getenv("APP_SECRET_KEY") or "cni-secret-key-2026").strip()

app = FastAPI(
    title="Sistema Indústria | People Analytics - Estrutura Organizacional",
    version="1.3.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

_HEADCOUNT_CACHE: dict[str, tuple[list, str]] = {}
_CUSTOM_UPLOADED_DATA: Optional[tuple[list, str]] = None


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


def get_repository() -> tuple[IHeadcountRepository, str]:
    db_repo = DatabricksHeadcountRepository()
    if db_repo.is_configured():
        return db_repo, "Databricks Unity Catalog"

    env_excel = (os.getenv("FALLBACK_EXCEL_PATH") or "").strip()
    excel_path = Path(env_excel) if env_excel else DEFAULT_EXCEL_PATH
    if not excel_path.is_absolute():
        excel_path = PROJECT_ROOT / excel_path

    if not excel_path.exists():
        raise FileNotFoundError(
            f"Arquivo padrão não localizado em '{excel_path}'. "
            f"Utilize o botão 'Subir Base (.xlsx/.csv)' no topo da página."
        )

    return ExcelHeadcountRepository(str(excel_path), excel_path.name), f"Base Padrão ({excel_path.name})"


def fetch_cached_records(reference_date: str, force_refresh: bool = False) -> tuple[list, str]:
    global _CUSTOM_UPLOADED_DATA
    if _CUSTOM_UPLOADED_DATA is not None and not force_refresh:
        return _CUSTOM_UPLOADED_DATA

    if force_refresh or reference_date not in _HEADCOUNT_CACHE:
        try:
            repo, source_label = get_repository()
            records = repo.fetch_current_headcount(reference_date)
            _HEADCOUNT_CACHE[reference_date] = (records, source_label)
            if force_refresh:
                _CUSTOM_UPLOADED_DATA = None
        except Exception as err:
            if DEFAULT_EXCEL_PATH.exists():
                fallback_repo = ExcelHeadcountRepository(str(DEFAULT_EXCEL_PATH), DEFAULT_EXCEL_PATH.name)
                records = fallback_repo.fetch_current_headcount(reference_date)
                _HEADCOUNT_CACHE[reference_date] = (records, f"Fallback Local ({DEFAULT_EXCEL_PATH.name})")
            else:
                raise err

    return _HEADCOUNT_CACHE[reference_date]


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
        return {
            "status": "authenticated",
            "user": user_record["username"],
            "role": user_record["role"],
        }
    raise HTTPException(status_code=401, detail="Utilizador ou palavra-passe inválidos.")


@app.post("/api/logout")
async def logout(response: Response):
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
    reference_date: str = Form("2026-10-01"),
):
    global _CUSTOM_UPLOADED_DATA
    require_auth(request, allow_viewer=False)

    fname = (data_file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xls") or fname.endswith(".csv")):
        raise HTTPException(
            status_code=400,
            detail="Formato inválido. Envie um arquivo .xlsx ou .csv extraído do Databricks.",
        )

    try:
        raw_bytes = await data_file.read()
        repo = ExcelHeadcountRepository(raw_bytes, filename=data_file.filename or "upload.xlsx")
        records = repo.fetch_current_headcount(reference_date)
        source_label = f"Upload Web ({data_file.filename})"
        _CUSTOM_UPLOADED_DATA = (records, source_label)
        _HEADCOUNT_CACHE[reference_date] = (records, source_label)

        summary = PeopleAnalyticsService.build_organization_summary(records=records)
        return JSONResponse(
            content={
                "status": "success",
                "source": source_label,
                "total_records": len(records),
                "total_headcount": summary["total_headcount"],
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
    ref_date = reference_date or os.getenv("DEFAULT_REFERENCE_DATE", "2026-10-01")
    try:
        records, source_label = fetch_cached_records(ref_date, force_refresh=refresh)
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
                "reference_date": ref_date,
                "default_label": os.getenv("DEFAULT_REFERENCE_LABEL", "Base Outubro 2026"),
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
    reference_date: str = Form("2026-10-01"),
    reference_label: str = Form("Base Outubro 2026"),
    custom_overrides_json: str = Form("[]"),
    auto_layout: bool = Form(True),
):
    require_auth(request, allow_viewer=False)
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Envie um arquivo válido com extensão .pptx")

    try:
        raw_bytes = await file.read()
        pptx_stream = io.BytesIO(raw_bytes)
        custom_overrides = json.loads(custom_overrides_json) if custom_overrides_json else []

        if data_file and data_file.filename:
            data_bytes = await data_file.read()
            temp_repo = ExcelHeadcountRepository(data_bytes, filename=data_file.filename)
            records = temp_repo.fetch_current_headcount(reference_date)
        else:
            records, _ = fetch_cached_records(reference_date)

        if not MAPPING_JSON_PATH.exists():
            raise FileNotFoundError(f"Arquivo de mapeamento não encontrado: {MAPPING_JSON_PATH}")

        updater = PresentationUpdaterService(mapping_json_path=str(MAPPING_JSON_PATH))
        updated_pptx_bytes, report = updater.execute_update(
            pptx_stream=pptx_stream,
            records=records,
            reference_label=reference_label,
            original_filename=file.filename,
            custom_overrides=custom_overrides,
            auto_layout=auto_layout,
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
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Erro interno no processamento do PowerPoint: {str(exc)}")