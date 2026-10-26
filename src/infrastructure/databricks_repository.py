import os
import io
import sys
import logging
from typing import List
import pandas as pd
from src.domain.models import HeadcountRecord
from src.domain.interfaces import IHeadcountRepository

# Silencia o warning [WARN] pyarrow is not installed by default since databricks-sql-connector 4.0.0
logging.getLogger("databricks.sql").setLevel(logging.ERROR)


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

        query = f"""
            SELECT 
                SEGMENTO_DES,
                AREA_DES,
                UNIDADE_DES,
                COUNT(CODIGO_CONTRATO) AS QTD_CONTRATOS
            FROM {self.catalog_table}
            WHERE DATA_RESCISAO IS NULL
              AND ANO_MES_REFERENCIA = %(ref_date)s
              AND VINCULO_EMPREGATICIO <> 'Estagiario'
            GROUP BY SEGMENTO_DES, AREA_DES, UNIDADE_DES
            ORDER BY SEGMENTO_DES, AREA_DES, UNIDADE_DES
        """

        records: List[HeadcountRecord] = []
        with self._build_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(query, {"ref_date": reference_date})
                rows = cursor.fetchall()
                for row in rows:
                    records.append(
                        HeadcountRecord(
                            segmento_des=str(row.SEGMENTO_DES).strip(),
                            area_des=str(row.AREA_DES).strip(),
                            unidade_des=str(row.UNIDADE_DES).strip(),
                            headcount=int(row.QTD_CONTRATOS),
                            ano_mes_referencia=reference_date,
                        )
                    )
        return records


class ExcelHeadcountRepository(IHeadcountRepository):
    """
    Repositório resiliente para leitura direta de extração Excel (.xlsx) ou CSV (.csv)
    tanto a partir de caminho em disco quanto via upload em memória (Bytes).
    """

    def __init__(self, file_source: str | bytes, filename: str = "dados.xlsx") -> None:
        self.file_source = file_source
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
                "SEGMENTO_DES, AREA_DES, UNIDADE_DES e a contagem de contratos."
            )

        col_seg = df.columns[0]
        col_area = df.columns[1]
        col_unit = df.columns[2]
        col_count = df.columns[3]

        records: List[HeadcountRecord] = []
        for _, row in df.iterrows():
            if pd.isna(row[col_seg]) or pd.isna(row[col_unit]):
                continue
            raw_count = pd.to_numeric(row[col_count], errors="coerce")
            if pd.isna(raw_count):
                continue
            records.append(
                HeadcountRecord(
                    segmento_des=str(row[col_seg]).strip(),
                    area_des=str(row[col_area]).strip() if not pd.isna(row[col_area]) else "",
                    unidade_des=str(row[col_unit]).strip(),
                    headcount=int(round(float(raw_count))),
                    ano_mes_referencia=reference_date,
                )
            )
        return records