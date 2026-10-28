from typing import List, Dict, Any, Optional
from src.domain.models import HeadcountRecord, OrganizationalRole


class PeopleAnalyticsService:
    """
    Serviço 100% dinâmico de inteligência organizacional:
    - Não depende de slide_mapping.json nem de nomes de áreas fixos no código;
    - Determina se um agrupamento (AREA_DES) é Superintendência ou Staff Direto/Diretoria
      verificando dinamicamente se existe titular de Superintendência dentro do agrupamento.
    """

    @classmethod
    def build_organization_summary(
        cls,
        records: List[HeadcountRecord],
        segmento_filter: Optional[str] = None,
        area_filter: Optional[str] = None,
        search_term: Optional[str] = None,
    ) -> Dict[str, Any]:
        active_records = [
            HeadcountRecord(
                segmento_des=str(r.segmento_des or "").strip(),
                area_des=str(r.area_des or "").strip(),
                unidade_des=str(r.unidade_des or "").strip(),
                headcount=r.headcount,
                ano_mes_referencia=r.ano_mes_referencia,
            )
            for r in records
            if r.headcount > 0
        ]

        # Identifica dinamicamente quais (segmento, area) possuem um titular de Superintendência
        areas_with_superintendencia = {
            (r.segmento_des, r.area_des)
            for r in active_records
            if r.organizational_role == OrganizationalRole.SUPERINTENDENCIA_TITULAR
        }

        filtered = active_records
        if segmento_filter and segmento_filter != "TODOS":
            filtered = [r for r in filtered if r.segmento_des == segmento_filter.strip()]
        if area_filter and area_filter != "TODAS":
            filtered = [r for r in filtered if r.area_des == area_filter.strip()]
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
            seg = by_segment.setdefault(
                r.segmento_des, {"total": 0, "slide_number": 1, "areas": {}}
            )
            seg["total"] += r.headcount

            # Uma área é Staff Direto / Diretoria se NÃO for encabeçada por uma Superintendência Titular
            is_staff = (r.segmento_des, r.area_des) not in areas_with_superintendencia

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
                    "slide_number": 1,
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
            "available_segments": sorted({r.segmento_des for r in active_records}),
            "available_areas": sorted(
                {
                    r.area_des
                    for r in active_records
                    if not segmento_filter or segmento_filter == "TODOS" or r.segmento_des == segmento_filter.strip()
                }
            ),
            "hierarchy": sorted_segments,
        }