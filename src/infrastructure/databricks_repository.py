import os
import io
from typing import List
import pandas as pd
from databricks import sql
from src.domain.models import HeadcountRecord
from src.domain.interfaces import IHeadcountRepository


class DatabricksHeadcountRepository(IHeadcountRepository):
    """
    Conector multi-estratégia para o Databricks Unity Catalog.
    Suporta:
    1. Token PAT (DATABRICKS_TOKEN)
    2. OAuth M2M para Vercel (DATABRICKS_CLIENT_ID + DATABRICKS_CLIENT_SECRET)
    3. OAuth U2M SSO via Navegador sem Token (DATABRICKS_USE_SSO=true no VS Code local)
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

    def is_configured(self) -> bool:
        if not (self.server_hostname and self.http_path):
            return False
        return bool(
            self.access_token
            or (self.client_id and self.client_secret)
            or self.use_sso
        )

    def _build_connection(self):
        if self.access_token:
            return sql.connect(
                server_hostname=self.server_hostname,
                http_path=self.http_path,
                access_token=self.access_token,
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
            )

        if self.use_sso:
            return sql.connect(
                server_hostname=self.server_hostname,
                http_path=self.http_path,
                auth_type="databricks-oauth",
            )

        raise EnvironmentError("Nenhum método de autenticação Databricks configurado no .env.")

    def fetch_current_headcount(self, reference_date: str = "2026-10-01") -> List[HeadcountRecord]:
        if not self.is_configured():
            raise EnvironmentError(
                "Configuração do Databricks incompleta no .env. "
                "Informe DATABRICKS_SERVER_HOSTNAME e DATABRICKS_HTTP_PATH."
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
    Repositório para leitura direta de extração Excel (.xlsx) ou CSV (.csv)
    tanto a partir de caminho em disco quanto via upload em memória (Bytes).
    """

    def __init__(self, file_source: str | bytes, filename: str = "dados.xlsx") -> None:
        self.file_source = file_source
        self.filename = filename.lower()

    def fetch_current_headcount(self, reference_date: str = "2026-10-01") -> List[HeadcountRecord]:
        if isinstance(self.file_source, bytes):
            stream = io.BytesIO(self.file_source)
            if self.filename.endswith(".csv"):
                try:
                    df = pd.read_csv(stream, sep=None, engine="python")
                except Exception:
                    stream.seek(0)
                    df = pd.read_csv(stream)
            else:
                xls = pd.ExcelFile(stream)
                df = pd.read_excel(stream, sheet_name=xls.sheet_names[0])
        else:
            if str(self.file_source).lower().endswith(".csv"):
                df = pd.read_csv(self.file_source, sep=None, engine="python")
            else:
                xls = pd.ExcelFile(self.file_source)
                df = pd.read_excel(self.file_source, sheet_name=xls.sheet_names[0])

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
            records.append(
                HeadcountRecord(
                    segmento_des=str(row[col_seg]).strip(),
                    area_des=str(row[col_area]).strip(),
                    unidade_des=str(row[col_unit]).strip(),
                    headcount=int(float(row[col_count])),
                    ano_mes_referencia=reference_date,
                )
            )
        return records