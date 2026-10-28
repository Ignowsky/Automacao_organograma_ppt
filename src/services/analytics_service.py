from typing import List, Dict, Any, Optional
from src.domain.models import HeadcountRecord, OrganizationalRole
from src.infrastructure.pptx_engine import normalize_text_key


class PeopleAnalyticsService:
    """
    Serviço 100% Dinâmico de Inteligência Organizacional para o Mapa de Headcount:
    - Não depende de config/slide_mapping.json;
    - Associa automaticamente cada Diretoria/Entidade ao seu respectivo número de slide (1 a 11);
    - Classifica papéis hierárquicos pela Gramática Universal de 5 Níveis.
    """

    @staticmethod
    def _infer_slide_number(segmento: str) -> int:
        s_norm = normalize_text_key(segmento)
        if "governanca" in s_norm:
            return 2
        if "juridica" in s_norm:
            return 3
        if "relacoes institucionais" in s_norm:
            return 4
        if "desenvolvimento industrial" in s_norm:
            return 5
        if "comunicacao" in s_norm:
            return 6
        if "corporativa" in s_norm:
            return 7
        if s_norm == "sesi" or "sesi" in s_norm:
            return 8
        if s_norm == "senai" or "senai" in s_norm:
            return 9
        if s_norm == "iel" or "iel" in s_norm:
            return 10
        if "compliance" in s_norm:
            return 11
        return 1

    @staticmethod
    def _normalize_segmento(segmento: str, unidade: str) -> str:
        seg_clean = str(segmento or "").strip()
        uni_clean = str(unidade or "").strip()
        if "Compliance" in seg_clean or "Compliance" in uni_clean:
            return "Diretoria de Compliance e Integridade"
        return seg_clean

    @classmethod
    def build_organization_summary(
        cls,
        records: List[HeadcountRecord],
        segmento_filter: Optional[str] = None,
        area_filter: Optional[str] = None,
        search_term: Optional[str] = None,
    ) -> Dict[str, Any]:
        normalized_records = [
            HeadcountRecord(
                segmento_des=cls._normalize_segmento(r.segmento_des, r.unidade_des),
                area_des=r.area_des,
                unidade_des=r.unidade_des,
                headcount=r.headcount,
                ano_mes_referencia=r.ano_mes_referencia,
            )
            for r in records
            if r.headcount > 0
        ]

        filtered = normalized_records
        if segmento_filter and segmento_filter != "TODOS":
            norm_seg_filter = cls._normalize_segmento(segmento_filter, "")
            filtered = [r for r in filtered if r.segmento_des == norm_seg_filter]
        if area_filter and area_filter != "TODAS":
            filtered = [r for r in filtered if r.area_des == area_filter]
        if search_term:
            term = search_term.lower().strip()
            filtered = [
                r for r in filtered
                if term in r.unidade_des.lower()
                or term in r.area_des.lower()
                or term in r.segmento_des.lower()
            ]

        total_headcount = sum(r.headcount for r in filtered)
        by_segment: Dict[str, Dict[str, Any]] = {}

        for r in filtered:
            role = r.organizational_role.value
            slide_num = cls._infer_slide_number(r.segmento_des)
            seg = by_segment.setdefault(
                r.segmento_des, {"total": 0, "slide_number": slide_num, "areas": {}}
            )
            seg["total"] += r.headcount

            is_staff = r.is_direct_staff_or_board

            area = seg["areas"].setdefault(
                r.area_des,
                {
                    "subtotal": 0,
                    "is_direct_staff": is_staff,
                    "unidades": [],
                },
            )
            area["subtotal"] += r.headcount
            area["unidades"].append(
                {
                    "unidade_des": r.unidade_des,
                    "original_unidade_des": r.unidade_des,
                    "headcount": r.headcount,
                    "role": role,
                    "manager": "",
                    "shape_key": "",
                    "slide_number": slide_num,
                }
            )

        role_priority = {
            OrganizationalRole.CONSELHO_SUPERIOR.value: 0,
            OrganizationalRole.DIRETORIA_CABECA.value: 1,
            OrganizationalRole.DIRETORIA_VINCULADA.value: 2,
            OrganizationalRole.SUPERINTENDENCIA_TITULAR.value: 3,
            OrganizationalRole.GERENCIA_UNIDADE.value: 4,
        }

        for seg_name, seg_obj in by_segment.items():
            for area_name, area_obj in seg_obj["areas"].items():
                area_obj["unidades"].sort(
                    key=lambda u: (role_priority.get(u["role"], 9), -u["headcount"], u["unidade_des"])
                )

            sorted_areas = dict(
                sorted(
                    seg_obj["areas"].items(),
                    key=lambda item: (0 if item[1]["is_direct_staff"] else 1, -item[1]["subtotal"]),
                )
            )
            seg_obj["areas"] = sorted_areas

        sorted_segments = dict(
            sorted(by_segment.items(), key=lambda item: -item[1]["total"])
        )

        return {
            "total_headcount": total_headcount,
            "total_unidades": len(filtered),
            "available_segments": sorted({r.segmento_des for r in normalized_records}),
            "available_areas": sorted(
                {
                    r.area_des
                    for r in normalized_records
                    if not segmento_filter or segmento_filter == "TODOS" or r.segmento_des == cls._normalize_segmento(segmento_filter, "")
                }
            ),
            "hierarchy": sorted_segments,
        }