import os
import io
import re
import sys
import logging
from datetime import datetime, date
from typing import List, Optional, Any
import pandas as pd
from src.domain.models import HeadcountRecord
from src.domain.interfaces import IHeadcountRepository

# Silencia o warning [WARN] pyarrow is not installed by default since databricks-sql-connector 4.0.0
logging.getLogger("databricks.sql").setLevel(logging.ERROR)

PT_MONTHS = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

PT_MONTHS_REVERSE = {
    "janeiro": 1, "jan": 1,
    "fevereiro": 2, "fev": 2,
    "marco": 3, "março": 3, "mar": 3,
    "abril": 4, "abr": 4,
    "maio": 5, "mai": 5,
    "junho": 6, "jun": 6,
    "julho": 7, "jul": 7,
    "agosto": 8, "ago": 8,
    "setembro": 9, "set": 9,
    "outubro": 10, "out": 10,
    "novembro": 11, "nov": 11,
    "dezembro": 12, "dez": 12,
}


def try_parse_any_date(raw_value: Any) -> Optional[tuple[int, int]]:
    """
    Tenta extrair (year, month) de qualquer formato de data/hora:
    - '2026-05-01T00:00:00.000+00:00' (Databricks ISO 8601 com timezone)
    - '2026-05-01 00:00:00' / '2026-05-01'
    - '05/2026' / '5/2026'
    - '01/05/2026' (dd/MM/yyyy)
    - Objetos datetime, date ou pd.Timestamp do Excel
    - Textos como 'Base Maio 2026' ou 'Maio/2026'
    Retorna None se o valor não for uma data válida.
    """
    if raw_value is None or pd.isna(raw_value):
        return None

    if isinstance(raw_value, (datetime, date, pd.Timestamp)):
        return int(raw_value.year), int(raw_value.month)

    if isinstance(raw_value, (int, float)):
        return None

    s = str(raw_value).strip()
    if not s:
        return None

    # 1. ISO 8601 completo ou parcial: '2026-05-01T00:00:00.000+00:00', '2026-05-01', '2026/05/01'
    m_iso = re.match(r"^(\d{4})[/\-](\d{1,2})(?:[/\-](\d{1,2}))?(?:[T\s].*)?$", s)
    if m_iso:
        y, m = int(m_iso.group(1)), int(m_iso.group(2))
        if 1990 <= y <= 2100 and 1 <= m <= 12:
            return y, m

    # 2. Formato MM/yyyy ou MM-yyyy: '05/2026', '5/2026'
    m_my = re.match(r"^(\d{1,2})[/\-](\d{4})$", s)
    if m_my:
        m, y = int(m_my.group(1)), int(m_my.group(2))
        if 1990 <= y <= 2100 and 1 <= m <= 12:
            return y, m

    # 3. Formato brasileiro dd/MM/yyyy (com ou sem hora): '01/05/2026'
    m_dmy = re.match(r"^(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})(?:[T\s].*)?$", s)
    if m_dmy:
        d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        if 1990 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31:
            return y, m

    # 4. Nome do mês em português: 'Base Maio 2026', 'Maio/2026', 'Maio_2026'
    s_low = s.lower()
    for m_name, m_num in sorted(PT_MONTHS_REVERSE.items(), key=lambda x: -len(x[0])):
        if re.search(rf"(?:^|[\s_\-/]){re.escape(m_name)}(?:$|[\s_\-/])", s_low):
            y_match = re.search(r"(20\d{2})", s_low)
            if y_match:
                return int(y_match.group(1)), m_num

    # 5. Tentativa via parser ISO/UTC do pandas apenas para strings com cara de timestamp
    if any(ch in s for ch in ("-", "/", "T")) and re.search(r"20\d{2}", s):
        try:
            dt = pd.to_datetime(s, utc=True, errors="raise")
            if 1990 <= dt.year <= 2100:
                return int(dt.year), int(dt.month)
        except Exception:
            pass

    return None


def parse_reference_date_info(
    raw_value: Optional[Any] = None,
    fallback_filename: str = "",
) -> dict[str, str]:
    """
    Formata a data detectada em:
    - iso_date: '2026-05-01'
    - ano_mes: '05/2026'
    - label: 'Base Maio 2026'
    """
    parsed = try_parse_any_date(raw_value)

    # Se não veio valor na célula, tenta extrair SOMENTE se o nome do arquivo tiver mês explícito em português
    # (Ignora nomes automáticos de download do Databricks como 'New_Query_2026_10_26...')
    if parsed is None and fallback_filename:
        fn_low = fallback_filename.lower()
        if not fn_low.startswith("new_query"):
            parsed = try_parse_any_date(fallback_filename)

    if parsed is None:
        return {
            "iso_date": "",
            "ano_mes": "--/----",
            "label": "Base Não Identificada",
        }

    year, month = parsed
    return {
        "iso_date": f"{year:04d}-{month:02d}-01",
        "ano_mes": f"{month:02d}/{year:04d}",
        "label": f"Base {PT_MONTHS[month]} {year:04d}",
    }


class DatabricksHeadcountRepository(IHeadcountRepository):
    """
    Conector multi-estratégia para o Databricks Unity Catalog.
    """

    def __init__(self) -> None:
        self.server_hostname = (os.getenv("DATABRICKS_SERVER_HOSTNAME") or "").strip()
        self.http_path = (os.getenv("DATABRICKS_HTTP_PATH") or "").strip()
        self.access_token = (os.getenv("DATABRICKS_TOKEN") or "").strip()
        self.client_id = (os.getenv("DATABRICKS_CLIENT_ID") or "").strip()
        self.client_secret = (os.getenv("DATABRICKS_CLIENT_SECRET") or "").strip()
        self.use_sso = (os.getenv("DATABRICKS_USE_SSO") or "false").strip().lower() == "true"
        self.catalog_table = (
            os.getenv("DATABRICKS_CATALOG_TABLE") or "datalake__biz.sdh.dh_gestaorh__extenso"
        ).strip()

    @staticmethod
    def _is_headless_serverless() -> bool:
        if os.getenv("VERCEL") or os.getenv("VERCEL_ENV") or os.getenv("AWS_LAMBDA_FUNCTION_NAME"):
            return True
        if sys.platform.startswith("linux") and not os.getenv("DISPLAY") and not os.getenv("WAYLAND_DISPLAY"):
            return True
        return False

    def is_configured(self) -> bool:
        if not (self.server_hostname and self.http_path):
            return False
        if self.access_token or (self.client_id and self.client_secret):
            return True
        if self.use_sso and not self._is_headless_serverless():
            return True
        return False

    def _build_connection(self):
        from databricks import sql

        if self.access_token:
            return sql.connect(
                server_hostname=self.server_hostname,
                http_path=self.http_path,
                access_token=self.access_token,
                _socket_timeout=20,
            )

        if self.client_id and self.client_secret:
            from databricks.sdk.core import Config, oauth_service_principal

            def credential_provider():
                config = Config(
                    host=f"https://{self.server_hostname}",
                    client_id=self.client_id,
                    client_secret=self.client_secret,
                )
                return oauth_service_principal(config)

            return sql.connect(
                server_hostname=self.server_hostname,
                http_path=self.http_path,
                credentials_provider=credential_provider,
                _socket_timeout=20,
            )

        if self.use_sso:
            if self._is_headless_serverless():
                raise EnvironmentError(
                    "SSO via navegador (U2M) não é suportado em ambiente Serverless (Vercel). "
                    "Utilize o upload de planilha (.xlsx/.csv)."
                )
            return sql.connect(
                server_hostname=self.server_hostname,
                http_path=self.http_path,
                auth_type="databricks-oauth",
            )

        raise EnvironmentError("Nenhum método de autenticação Databricks compatível configurado.")

    def fetch_current_headcount(self, reference_date: str = "") -> List[HeadcountRecord]:
        if not self.is_configured():
            raise EnvironmentError(
                "Databricks não configurado para este ambiente. Usando base Excel/CSV."
            )

        date_info = parse_reference_date_info(reference_date)
        sql_ref_date = date_info["iso_date"]

        # Se o utilizador informou uma data válida, filtra por ela; caso contrário, busca a mais recente do banco
        if sql_ref_date:
            where_date_clause = "AND ANO_MES_REFERENCIA = %(ref_date)s"
            params = {"ref_date": sql_ref_date}
        else:
            where_date_clause = (
                f"AND ANO_MES_REFERENCIA = ("
                f"SELECT MAX(ANO_MES_REFERENCIA) FROM {self.catalog_table} WHERE DATA_RESCISAO IS NULL"
                f")"
            )
            params = {}

        query = f"""
            SELECT 
                TO_CHAR(date_trunc('month', ANO_MES_REFERENCIA), 'MM/yyyy') AS Ano_Mes,
                SEGMENTO_DES,
                AREA_DES,
                UNIDADE_DES,
                COUNT(CODIGO_CONTRATO) AS QTD_CONTRATOS
            FROM {self.catalog_table}
            WHERE DATA_RESCISAO IS NULL
              {where_date_clause}
              AND VINCULO_EMPREGATICIO <> 'Estagiario'
            GROUP BY 
                ANO_MES_REFERENCIA,
                SEGMENTO_DES,
                AREA_DES,
                UNIDADE_DES
            ORDER BY SEGMENTO_DES, AREA_DES, UNIDADE_DES
        """

        records: List[HeadcountRecord] = []
        with self._build_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(query, params)
                rows = cursor.fetchall()
                for row in rows:
                    raw_dt = getattr(row, "Ano_Mes", None) or getattr(row, "ANO_MES_REFERENCIA", None)
                    row_info = parse_reference_date_info(raw_dt)
                    records.append(
                        HeadcountRecord(
                            segmento_des=str(row.SEGMENTO_DES).strip(),
                            area_des=str(row.AREA_DES).strip(),
                            unidade_des=str(row.UNIDADE_DES).strip(),
                            headcount=int(row.QTD_CONTRATOS),
                            ano_mes_referencia=row_info["ano_mes"],
                        )
                    )
        return records


class ExcelHeadcountRepository(IHeadcountRepository):
    """
    Repositório inteligente para leitura de planilhas (.xlsx ou .csv).
    Inspeciona tanto os cabeçalhos quanto os valores reais das células para localizar
    automaticamente a coluna de data (mesmo que venha como '2026-05-01T00:00:00.000+00:00',
    '05/2026', '2026-05-01' ou Timestamp do Excel, em qualquer posição da planilha).
    """

    def __init__(self, file_source: str | bytes, filename: str = "dados.xlsx") -> None:
        self.file_source = file_source
        self.raw_filename = filename.strip()
        self.filename = filename.lower().strip()

    @staticmethod
    def _detect_columns(df: pd.DataFrame) -> tuple[Optional[Any], Any, Any, Any, Any]:
        cols = list(df.columns)
        col_date = None

        # 1º Passo: Inspeciona os valores reais das células de cada coluna para achar a coluna de data!
        for c in cols:
            non_null_sample = df[c].dropna().head(10).tolist()
            if not non_null_sample:
                continue
            date_matches = sum(1 for val in non_null_sample if try_parse_any_date(val) is not None)
            if date_matches >= max(1, len(non_null_sample) // 2):
                col_date = c
                break

        # 2º Passo: Se não achou pelo valor, verifica pelo nome do cabeçalho
        if col_date is None:
            for c in cols:
                c_up = str(c).strip().upper()
                if any(k in c_up for k in ("ANO_MES", "REFERENCIA", "DATA", "COMPETENCIA", "DATE_TRUNC", "TO_CHAR")):
                    col_date = c
                    break

        remaining_cols = [c for c in cols if c != col_date]
        if len(remaining_cols) < 4:
            raise ValueError(
                "A planilha deve conter as colunas SEGMENTO_DES, AREA_DES, UNIDADE_DES e a contagem de contratos."
            )

        col_seg = None
        col_area = None
        col_unit = None
        col_count = None

        for c in remaining_cols:
            c_up = str(c).strip().upper()
            if col_seg is None and "SEGMENTO" in c_up:
                col_seg = c
            elif col_area is None and ("AREA" in c_up or "ÁREA" in c_up):
                col_area = c
            elif col_unit is None and "UNIDADE" in c_up:
                col_unit = c
            elif col_count is None and any(k in c_up for k in ("COUNT", "QTD", "CONTRATO", "HEADCOUNT", "TOTAL")):
                col_count = c

        # Se a coluna numérica de contagem ainda não foi identificada pelo nome, detecta pelo tipo numérico
        if col_count is None:
            for c in reversed(remaining_cols):
                if c in (col_seg, col_area, col_unit):
                    continue
                sample_num = pd.to_numeric(df[c].dropna().head(10), errors="coerce")
                if sample_num.notna().sum() >= 1:
                    col_count = c
                    break

        # Preenche eventuais colunas restantes pela ordem posicional
        unused = [c for c in remaining_cols if c not in (col_seg, col_area, col_unit, col_count)]
        if col_seg is None and unused:
            col_seg = unused.pop(0)
        if col_area is None and unused:
            col_area = unused.pop(0)
        if col_unit is None and unused:
            col_unit = unused.pop(0)
        if col_count is None and unused:
            col_count = unused.pop(0)

        return col_date, col_seg, col_area, col_unit, col_count

    def fetch_current_headcount(self, reference_date: str = "") -> List[HeadcountRecord]:
        if isinstance(self.file_source, bytes):
            stream = io.BytesIO(self.file_source)
            if self.filename.endswith(".csv"):
                try:
                    df = pd.read_csv(stream, sep=None, engine="python", encoding="utf-8-sig")
                except Exception:
                    stream.seek(0)
                    df = pd.read_csv(stream, sep=";", encoding="latin1")
            else:
                stream.seek(0)
                df = pd.read_excel(stream, sheet_name=0, engine="openpyxl")
        else:
            if str(self.file_source).lower().endswith(".csv"):
                try:
                    df = pd.read_csv(self.file_source, sep=None, engine="python", encoding="utf-8-sig")
                except Exception:
                    df = pd.read_csv(self.file_source, sep=";", encoding="latin1")
            else:
                df = pd.read_excel(self.file_source, sheet_name=0, engine="openpyxl")

        df = df.dropna(how="all")
        col_date, col_seg, col_area, col_unit, col_count = self._detect_columns(df)

        # Descobre a data da planilha varrendo a coluna de data (se existir)
        detected_sheet_ano_mes = None
        if col_date is not None:
            for val in df[col_date].dropna():
                parsed = try_parse_any_date(val)
                if parsed is not None:
                    detected_sheet_ano_mes = f"{parsed[1]:02d}/{parsed[0]:04d}"
                    break

        if not detected_sheet_ano_mes:
            fallback_info = parse_reference_date_info(raw_value=None, fallback_filename=self.raw_filename)
            detected_sheet_ano_mes = fallback_info["ano_mes"]

        records: List[HeadcountRecord] = []
        for _, row in df.iterrows():
            if pd.isna(row[col_seg]) or pd.isna(row[col_unit]):
                continue
            raw_count = pd.to_numeric(row[col_count], errors="coerce")
            if pd.isna(raw_count):
                continue

            row_ano_mes = detected_sheet_ano_mes
            if col_date is not None and not pd.isna(row[col_date]):
                parsed_row = try_parse_any_date(row[col_date])
                if parsed_row is not None:
                    row_ano_mes = f"{parsed_row[1]:02d}/{parsed_row[0]:04d}"

            records.append(
                HeadcountRecord(
                    segmento_des=str(row[col_seg]).strip(),
                    area_des=str(row[col_area]).strip() if not pd.isna(row[col_area]) else "",
                    unidade_des=str(row[col_unit]).strip(),
                    headcount=int(round(float(raw_count))),
                    ano_mes_referencia=row_ano_mes,
                )
            )
        return records