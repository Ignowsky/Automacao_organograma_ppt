import os
import json
from pathlib import Path
from typing import List, Dict, Any, Optional
from src.domain.models import HeadcountRecord, OrganizationalRole


class PeopleAnalyticsService:
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
        cls._load_mapping_metadata()

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
                or term in cls._MANAGER_LOOKUP.get((r.segmento_des, r.area_des, r.unidade_des), {}).get("manager", "").lower()
            ]

        total_headcount = sum(r.headcount for r in filtered)
        by_segment: Dict[str, Dict[str, Any]] = {}

        for r in filtered:
            meta = cls._MANAGER_LOOKUP.get((r.segmento_des, r.area_des, r.unidade_des), {})
            role = r.organizational_role.value
            seg = by_segment.setdefault(
                r.segmento_des, {"total": 0, "slide_number": meta.get("slide_number", 1), "areas": {}}
            )
            seg["total"] += r.headcount
            if meta.get("slide_number"):
                seg["slide_number"] = meta["slide_number"]

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
                    "unidade_des": meta.get("display_name") or r.unidade_des,
                    "original_unidade_des": r.unidade_des,
                    "headcount": r.headcount,
                    "role": role,
                    "manager": meta.get("manager", ""),
                    "shape_key": meta.get("shape_key", ""),
                    "slide_number": meta.get("slide_number", seg["slide_number"]),
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