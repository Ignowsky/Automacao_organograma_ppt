from typing import List, Protocol, BinaryIO
from src.domain.models import HeadcountRecord, ShapeMappingRule, ExecutionReport


class IHeadcountRepository(Protocol):
    """Contrato (DIP) para busca de dados (permite usar Databricks em Produção ou Excel/Mock em Dev)."""
    def fetch_current_headcount(self, reference_date: str) -> List[HeadcountRecord]:
        ...


class IPresentationUpdater(Protocol):
    """Contrato para inspeção, atualização cirúrgica e validação do arquivo PowerPoint em memória."""
    def execute_controlled_update(
        self,
        pptx_stream: BinaryIO,
        records: List[HeadcountRecord],
        rules: List[ShapeMappingRule],
        reference_label: str
    ) -> tuple[bytes, ExecutionReport]:
        ...