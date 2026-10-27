from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional, Dict, Any
import unicodedata


def _normalize_str(text: str) -> str:
    """Normaliza texto removendo acentos e espaços extras para classificação hierárquica."""
    if not text:
        return ""
    clean = " ".join(str(text).strip().split()).lower()
    nfkd = unicodedata.normalize("NFKD", clean)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


class UpdateStatus(str, Enum):
    UPDATED_NUMBER = "NUMERO_ATUALIZADO"
    UPDATED_NAME_AND_NUMBER = "NOME_E_NUMERO_ATUALIZADOS"
    CLONED_NEW_BOX = "NOVA_CAIXA_CRIADA_PADRAO_PPT"
    REMOVED_BOX = "CAIXA_REMOVIDA_PELO_UTILIZADOR"
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


class OrganizationalRole(str, Enum):
    """
    Gramática Hierárquica de 5 Níveis do Sistema Indústria (CNI, SESI, SENAI, IEL):
    Garante que Diretorias Vinculadas fiquem abaixo da Diretoria Geral e acima das Superintendências,
    e que Superintendências nunca sejam posicionadas no lugar de Gerências.
    """
    CONSELHO_SUPERIOR = "CONSELHO_SUPERIOR"               # Rank 0: Conselho Nacional SENAI, Presidência
    DIRETORIA_CABECA = "DIRETORIA_CABECA"                 # Rank 1: Diretoria Geral SENAI, Sup. Geral SESI, Gabinete Diretoria CNI
    DIRETORIA_VINCULADA = "DIRETORIA_VINCULADA"           # Rank 2: Diretoria interna (ex: Diretoria do Instituto Integrador no SENAI)
    SUPERINTENDENCIA_TITULAR = "SUPERINTENDENCIA_TITULAR" # Rank 3: Titular de Superintendência (ex: Sup. de Negócios Institucionais)
    GERENCIA_UNIDADE = "GERENCIA_UNIDADE"                 # Rank 4: Gerências subordinadas, Ouvidoria, Logística, etc.


@dataclass(frozen=True)
class HeadcountRecord:
    """Representa uma linha retornada pela query do Databricks ou planilha."""
    segmento_des: str
    area_des: str
    unidade_des: str
    headcount: int
    ano_mes_referencia: str = ""

    @property
    def organizational_role(self) -> OrganizationalRole:
        """Determina dinamicamente o papel hierárquico da unidade na árvore organizacional."""
        s_norm = _normalize_str(self.segmento_des)
        a_norm = _normalize_str(self.area_des)
        u_norm = _normalize_str(self.unidade_des)

        if u_norm.startswith("conselho") or u_norm.startswith("presidencia"):
            return OrganizationalRole.CONSELHO_SUPERIOR

        if (
            "diretoria geral" in u_norm
            or "superintendencia geral" in u_norm
            or u_norm == s_norm
            or (u_norm.startswith("diretoria") and a_norm == s_norm)
        ):
            return OrganizationalRole.DIRETORIA_CABECA

        if u_norm.startswith("diretoria") or a_norm.startswith("diretoria"):
            return OrganizationalRole.DIRETORIA_VINCULADA

        if u_norm.startswith("superintend"):
            return OrganizationalRole.SUPERINTENDENCIA_TITULAR

        return OrganizationalRole.GERENCIA_UNIDADE

    @property
    def hierarchy_rank(self) -> int:
        """Retorna a ordem numérica estrita na árvore (0 = topo até 4 = base operacional)."""
        rank_map = {
            OrganizationalRole.CONSELHO_SUPERIOR: 0,
            OrganizationalRole.DIRETORIA_CABECA: 1,
            OrganizationalRole.DIRETORIA_VINCULADA: 2,
            OrganizationalRole.SUPERINTENDENCIA_TITULAR: 3,
            OrganizationalRole.GERENCIA_UNIDADE: 4,
        }
        return rank_map.get(self.organizational_role, 4)

    @property
    def is_direct_staff_or_board(self) -> bool:
        """Indica se a área pertence ao topo diretivo/staff (acima das superintendências regulares)."""
        return (
            self.area_des == self.segmento_des
            or self.organizational_role in (
                OrganizationalRole.CONSELHO_SUPERIOR,
                OrganizationalRole.DIRETORIA_CABECA,
                OrganizationalRole.DIRETORIA_VINCULADA,
            )
            or _normalize_str(self.area_des) in ("cni/sp", "diretoria geral", "conselho nacional senai")
        )


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
    slides_preview: List[Dict[str, Any]] = field(default_factory=list)