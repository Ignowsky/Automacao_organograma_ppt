import os
import json
from pathlib import Path
from typing import List, Dict, Any, Optional
from src.domain.models import HeadcountRecord


class PeopleAnalyticsService:
    """
    Serviço de inteligência organizacional para o Painel Web:
    - Separa Gabinete/Titular de Superintendência das Gerências subordinadas.
    - Ordena Gerências por Headcount decrescente.
    - Enriquece cada unidade com o nome do Gestor mapeado e o Slide correspondente.
    """

    _MANAGER_LOOKUP: Dict[tuple[str, str, str], Dict[str, Any]] = {}

    @classmethod
    def _load_mapping_metadata(cls) -> None:
        if cls._MANAGER_LOOKUP:
            return
        mapping_path = Path(__file__).resolve().parent.parent.parent / "config" / "slide_mapping.json"
        if not mapping_path.exists() or os.path.getsize(mapping_path) == 0:
            return
        try:
            with open(mapping_path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            for u in data.get("unit_boxes", []):
                key = (u["segmento"], u["area"], u["unidade"])
                cls._MANAGER_LOOKUP[key] = {
                    "slide_number": int(u["slide_index"]) + 1,
                    "shape_key": f"{u['name_path']} | {u['badge_path']}",
                    "display_name": u.get("rename_to") or u["unidade"],
                    "manager": u.get("manager", ""),
                }
            for cl in data.get("cloned_new_boxes", []):
                key = (cl["segmento"], cl["area"], cl["unidade"])
                cls._MANAGER_LOOKUP[key] = {
                    "slide_number": int(cl["slide_index"]) + 1,
                    "shape_key": f"Clonado de {cl['template_group_id']}",
                    "display_name": cl["display_name"],
                    "manager": cl.get("manager_placeholder", "(Definir Gestor)"),
                }
        except Exception:
            pass

    @staticmethod
    def _classify_unit_role(area_name: str, unit_name: str) -> str:
        u_low = unit_name.lower().strip()
        if u_low.startswith("diretoria") or u_low.startswith("superintendência geral") or u_low.startswith("conselho"):
            return "DIRETORIA_CABECA"
        if u_low.startswith("superintend"):
            return "SUPERINTENDENCIA_TITULAR"
        return "GERENCIA_UNIDADE"

    @classmethod
    def build_organization_summary(
        cls,
        records: List[HeadcountRecord],
        segmento_filter: Optional[str] = None,
        area_filter: Optional[str] = None,
        search_term: Optional[str] = None,
    ) -> Dict[str, Any]:
        cls._load_mapping_metadata()

        filtered = records
        if segmento_filter and segmento_filter != "TODOS":
            filtered = [r for r in filtered if r.segmento_des == segmento_filter]
        if area_filter and area_filter != "TODAS":
            filtered = [r for r in filtered if r.area_des == area_filter]
        if search_term:
            term = search_term.lower().strip()
            filtered = [
                r for r in filtered
                if term in r.unidade_des.lower()
                or term in r.area_des.lower()
                or term in r.segmento_des.lower()
                or term in cls._MANAGER_LOOKUP.get((r.segmento_des, r.area_des, r.unidade_des), {}).get("manager", "").lower()
            ]

        total_headcount = sum(r.headcount for r in filtered)
        by_segment: Dict[str, Dict[str, Any]] = {}

        for r in filtered:
            meta = cls._MANAGER_LOOKUP.get((r.segmento_des, r.area_des, r.unidade_des), {})
            role = cls._classify_unit_role(r.area_des, r.unidade_des)
            seg = by_segment.setdefault(
                r.segmento_des, {"total": 0, "slide_number": meta.get("slide_number", 1), "areas": {}}
            )
            seg["total"] += r.headcount
            if meta.get("slide_number"):
                seg["slide_number"] = meta["slide_number"]

            area = seg["areas"].setdefault(
                r.area_des,
                {
                    "subtotal": 0,
                    "is_direct_staff": (r.area_des == r.segmento_des or r.area_des in ("CNI/SP", "Diretoria Geral", "Conselho Nacional SENAI")),
                    "unidades": [],
                },
            )
            area["subtotal"] += r.headcount
            area["unidades"].append(
                {
                    "unidade_des": meta.get("display_name") or r.unidade_des,
                    "original_unidade_des": r.unidade_des,
                    "headcount": r.headcount,
                    "role": role,
                    "manager": meta.get("manager", ""),
                    "shape_key": meta.get("shape_key", ""),
                    "slide_number": meta.get("slide_number", seg["slide_number"]),
                }
            )

        # Ordena áreas (Gabinete/Diretoria primeiro, depois Superintendências por maior subtotal)
        # E dentro de cada área: Titular da Superintendência no TOPO, seguido por Gerências em ordem decrescente de HC!
        role_priority = {"DIRETORIA_CABECA": 0, "SUPERINTENDENCIA_TITULAR": 1, "GERENCIA_UNIDADE": 2}
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

        # Ordena Segmentos por maior Headcount total
        sorted_segments = dict(
            sorted(by_segment.items(), key=lambda item: -item[1]["total"])
        )

        return {
            "total_headcount": total_headcount,
            "total_unidades": len(filtered),
            "available_segments": sorted({r.segmento_des for r in records}),
            "available_areas": sorted(
                {
                    r.area_des
                    for r in records
                    if not segmento_filter or segmento_filter == "TODOS" or r.segmento_des == segmento_filter
                }
            ),
            "hierarchy": sorted_segments,
        }