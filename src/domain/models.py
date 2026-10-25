from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class UpdateStatus(str, Enum):
    UPDATED_NUMBER = "NUMERO_ATUALIZADO"
    UPDATED_NAME_AND_NUMBER = "NOME_E_NUMERO_ATUALIZADOS"
    CLONED_NEW_BOX = "NOVA_CAIXA_CRIADA_PADRAO_PPT"
    EXTINCT_IN_DATABRICKS = "UNIDADE_SEM_CONTRATOS_NO_BANCO"
    UNCHANGED = "INALTERADO"
    ERROR = "ERRO_TECNICO"


class TargetType(str, Enum):
    HEADCOUNT_UNIT = "HEADCOUNT_UNIDADE"           # Ex: Gerência ou linha direta da Superintendência
    HEADCOUNT_AREA_SUBTOTAL = "SUBTOTAL_AREA"      # Ex: Badges no rodapé das colunas (Soma de AREA_DES)
    HEADCOUNT_SEGMENT_TOTAL = "TOTAL_DIRETORIA"    # Ex: Total do slide ou Macroestrutura (Soma de SEGMENTO_DES)
    HEADCOUNT_COMPANY_TOTAL = "TOTAL_GERAL_CNI"    # Ex: Shape 66 do Slide 0 (938)
    AREA_NAME = "NOME_DA_AREA"                     # Substitui o nome da área preservando o (Nome do Gestor)
    REFERENCE_DATE = "DATA_BASE_REFERENCIA"        # Ex: "Base Maio 2026" -> "Base Outubro 2026"


@dataclass(frozen=True)
class HeadcountRecord:
    """Representa uma linha retornada pela query do Databricks."""
    segmento_des: str
    area_des: str
    unidade_des: str
    headcount: int
    ano_mes_referencia: str = "2026-10-01"


@dataclass(frozen=True)
class ShapeMappingRule:
    """Regra explícita e auditável que vincula uma Shape do PPTX ao dado do Databricks."""
    slide_index: int
    shape_id: str                  # Suporta IDs aninhados em grupos, ex: "5.7" ou "138.140"
    target_type: TargetType
    segmento_key: Optional[str] = None
    area_key: Optional[str] = None
    unidade_key: Optional[str] = None
    expected_old_text: Optional[str] = None
    description: str = ""


@dataclass
class ChangeLogEntry:
    """Registro individual de auditoria para cada alteração planejada/executada."""
    slide_index: int
    shape_id: str
    shape_name: str
    target_type: TargetType
    old_value: str
    new_value: str
    status: UpdateStatus
    message: str = ""


@dataclass
class ExecutionReport:
    """Relatório final (Fase 9) retornado pela API e exibido na interface da Vercel."""
    original_filename: str
    output_filename: str
    slides_before: int
    slides_after: int
    total_company_headcount: int
    reference_label: str
    updated_count: int = 0
    unchanged_count: int = 0
    conflict_count: int = 0
    changes: List[ChangeLogEntry] = field(default_factory=list)
    technical_validation_passed: bool = False
    validation_notes: List[str] = field(default_factory=list)