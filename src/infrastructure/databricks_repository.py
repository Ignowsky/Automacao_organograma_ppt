import os
import io
import re
import sys
import logging
from datetime import datetime, date
from typing import List, Optional
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


def parse_reference_date_info(
    raw_value: Optional[ object ] = None,
    fallback_filename: str = "",
) -> dict[str, str]:
    """
    Converte qualquer representação de data ('05/2026', '2026-05-01', Timestamp ou nome de arquivo
    como 'Base_Maio_2026.xlsx') em um dicionário padronizado:
    - iso_date: '2026-05-01' (para filtro SQL ANO_MES_REFERENCIA)
    - ano_mes: '05/2026' (para exibição no campo REF e coluna Ano_Mes)
    - label: 'Base Maio 2026' (para tag dos 11 slides do PPTX e cabeçalho)
    """
    year: Optional[int] = None
    month: Optional[int] = None

    if raw_value is not None and not pd.isna(raw_value):
        if isinstance(raw_value, (datetime, date, pd.Timestamp)):
            year, month = raw_value.year, raw_value.month
        else:
            s = str(raw_value).strip()
            # 1) Formato MM/yyyy (ex: 05/2026)
            m_slash = re.match(r"^(\d{1,2})[/\-](\d{4})$", s)
            if m_slash:
                month, year = int(m_slash.group(1)), int(m_slash.group(2))
            else:
                # 2) Formato yyyy-MM-dd ou yyyy-MM (ex: 2026-05-01)
                m_iso = re.match(r"^(\d{4})[/\-](\d{1,2})(?:[/\-]\d{1,2})?", s)
                if m_iso:
                    year, month = int(m_iso.group(1)), int(m_iso.group(2))
                else:
                    # 3) Texto como "Base Maio 2026" ou "Maio/2026"
                    s_low = s.lower()
                    for m_name, m_num in PT_MONTHS_REVERSE.items():
                        if re.search(rf"\b{m_name}\b", s_low):
                            y_match = re.search(r"(20\d{2})", s_low)
                            if y_match:
                                month = m_num
                                year = int(y_match.group(1))
                                break

    # Fallback: tenta extrair mês e ano do nome do arquivo (ex: 'Base_Maio_2026.xlsx')
    if (year is None or month is None) and fallback_filename:
        fn_low = fallback_filename.lower()
        for m_name, m_num in sorted(PT_MONTHS_REVERSE.items(), key=lambda x: -len(x[0])):
            if m_name in fn_low:
                y_match = re.search(r"(20\d{2})", fn_low)
                if y_match:
                    month = m_num
                    year = int(y_match.group(1))
                    break
        if year is None or month is None:
            m_fn_iso = re.search(r"(20\d{2})[_\-](\d{2})", fn_low)
            if m_fn_iso:
                y_cand, m_cand = int(m_fn_iso.group(1)), int(m_fn_iso.group(2))
                if 1 <= m_cand <= 12:
                    year, month = y_cand, m_cand

    if year is None or month is None or not (1 <= month <= 12):
        year, month = 2026, 10

    iso_date = f"{year:04d}-{month:02d}-01"
    ano_mes = f"{month:02d}/{year:04d}"
    label = f"Base {PT_MONTHS[month]} {year:04d}"
    return {
        "iso_date": iso_date,
        "ano_mes": ano_mes,
        "label": label,
    }


class DatabricksHeadcountRepository(IHeadcountRepository):
    """
    Conector multi-estratégia para o Databricks Unity Catalog.
    Suporta:
    1. Token PAT (DATABRICKS_TOKEN)
    2. OAuth M2M para Vercel (DATABRICKS_CLIENT_ID + DATABRICKS_CLIENT_SECRET)
    3. OAuth U2M SSO via Navegador sem Token (DATABRICKS_USE_SSO=true APENAS em máquina local com tela)
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
        """
        Detecta se o código está rodando na Vercel / AWS Lambda ou Linux sem interface gráfica.
        Evita que o auth_type='databricks-oauth' trave o servidor por 300 segundos esperando navegador.
        """
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

    def fetch_current_headcount(self, reference_date: str = "2026-10-01") -> List[HeadcountRecord]:
        if not self.is_configured():
            raise EnvironmentError(
                "Databricks não configurado para este ambiente. Usando base Excel/CSV."
            )

        date_info = parse_reference_date_info(reference_date)
        sql_ref_date = date_info["iso_date"]

        query = f"""
            SELECT 
                TO_CHAR(date_trunc('month', ANO_MES_REFERENCIA), 'MM/yyyy') AS Ano_Mes,
                SEGMENTO_DES,
                AREA_DES,
                UNIDADE_DES,
                COUNT(CODIGO_CONTRATO) AS QTD_CONTRATOS
            FROM {self.catalog_table}
            WHERE DATA_RESCISAO IS NULL
              AND ANO_MES_REFERENCIA = %(ref_date)s
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
                cursor.execute(query, {"ref_date": sql_ref_date})
                rows = cursor.fetchall()
                for row in rows:
                    row_ano_mes = getattr(row, "Ano_Mes", None) or getattr(row, "ANO_MES", None) or date_info["ano_mes"]
                    records.append(
                        HeadcountRecord(
                            segmento_des=str(row.SEGMENTO_DES).strip(),
                            area_des=str(row.AREA_DES).strip(),
                            unidade_des=str(row.UNIDADE_DES).strip(),
                            headcount=int(row.QTD_CONTRATOS),
                            ano_mes_referencia=str(row_ano_mes).strip(),
                        )
                    )
        return records


class ExcelHeadcountRepository(IHeadcountRepository):
    """
    Repositório resiliente para leitura direta de extração Excel (.xlsx) ou CSV (.csv).
    Suporta dinamicamente:
    - O novo formato de 5 colunas: Ano_Mes, SEGMENTO_DES, AREA_DES, UNIDADE_DES, COUNT(CODIGO_CONTRATO)
    - O formato anterior de 4 colunas: SEGMENTO_DES, AREA_DES, UNIDADE_DES, COUNT(CODIGO_CONTRATO)
    """

    def __init__(self, file_source: str | bytes, filename: str = "dados.xlsx") -> None:
        self.file_source = file_source
        self.raw_filename = filename.strip()
        self.filename = filename.lower().strip()

    def fetch_current_headcount(self, reference_date: str = "2026-10-01") -> List[HeadcountRecord]:
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

        if len(df.columns) < 4:
            raise ValueError(
                "O arquivo enviado deve conter pelo menos 4 colunas: "
                "(Opcional: Ano_Mes), SEGMENTO_DES, AREA_DES, UNIDADE_DES e a contagem de contratos."
            )

        # Mapeamento inteligente de colunas por nome ou por posição (5 colunas vs 4 colunas)
        col_date = None
        col_seg = None
        col_area = None
        col_unit = None
        col_count = None

        for c in df.columns:
            c_up = str(c).strip().upper()
            if col_date is None and any(k in c_up for k in ("ANO_MES", "REFERENCIA", "DATA", "COMPETENCIA")):
                col_date = c
            elif col_seg is None and "SEGMENTO" in c_up:
                col_seg = c
            elif col_area is None and ("AREA" in c_up or "ÁREA" in c_up):
                col_area = c
            elif col_unit is None and "UNIDADE" in c_up:
                col_unit = c
            elif col_count is None and any(k in c_up for k in ("COUNT", "QTD", "CONTRATO", "HEADCOUNT", "TOTAL")):
                col_count = c

        # Fallback posicional caso os cabeçalhos estejam customizados
        if len(df.columns) >= 5:
            col_date = col_date if col_date is not None else df.columns[0]
            col_seg = col_seg if col_seg is not None else df.columns[1]
            col_area = col_area if col_area is not None else df.columns[2]
            col_unit = col_unit if col_unit is not None else df.columns[3]
            col_count = col_count if col_count is not None else df.columns[4]
        else:
            col_seg = col_seg if col_seg is not None else df.columns[0]
            col_area = col_area if col_area is not None else df.columns[1]
            col_unit = col_unit if col_unit is not None else df.columns[2]
            col_count = col_count if col_count is not None else df.columns[3]

        # Determina a competência padrão caso a planilha tenha apenas 4 colunas (inferindo pelo nome do arquivo se possível)
        fallback_info = parse_reference_date_info(
            raw_value=None,
            fallback_filename=self.raw_filename,
        )
        # Se o nome do arquivo não tinha mês/ano explícito, usa o reference_date informado
        if fallback_info["iso_date"] == "2026-10-01" and "outubro" not in self.filename and "2026_10" not in self.filename:
            fallback_info = parse_reference_date_info(reference_date, fallback_filename=self.raw_filename)

        records: List[HeadcountRecord] = []
        for _, row in df.iterrows():
            if pd.isna(row[col_seg]) or pd.isna(row[col_unit]):
                continue
            raw_count = pd.to_numeric(row[col_count], errors="coerce")
            if pd.isna(raw_count):
                continue

            if col_date is not None and not pd.isna(row[col_date]):
                row_date_info = parse_reference_date_info(row[col_date], fallback_filename=self.raw_filename)
                ref_ano_mes = row_date_info["ano_mes"]
            else:
                ref_ano_mes = fallback_info["ano_mes"]

            records.append(
                HeadcountRecord(
                    segmento_des=str(row[col_seg]).strip(),
                    area_des=str(row[col_area]).strip() if not pd.isna(row[col_area]) else "",
                    unidade_des=str(row[col_unit]).strip(),
                    headcount=int(round(float(raw_count))),
                    ano_mes_referencia=ref_ano_mes,
                )
            )
        return records