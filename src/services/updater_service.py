import io
import os
import json
from typing import List, Dict, Any, BinaryIO, Optional
from pptx import Presentation
from src.domain.models import (
    HeadcountRecord,
    ChangeLogEntry,
    ExecutionReport,
    UpdateStatus,
    TargetType,
)
from src.infrastructure.pptx_engine import PptxSurgicalEngine


class PresentationUpdaterService:
    """
    Orquestrador central das Fases 2 a 9 com suporte a Edições Manuais via Interface Web (Live Overrides).
    """

    def __init__(self, mapping_json_path: str = "config/slide_mapping.json") -> None:
        if not os.path.exists(mapping_json_path):
            raise FileNotFoundError(f"Arquivo de mapeamento não encontrado: {mapping_json_path}")
        if os.path.getsize(mapping_json_path) == 0:
            raise ValueError(f"O arquivo '{mapping_json_path}' está vazio (0 bytes).")

        with open(mapping_json_path, "r", encoding="utf-8-sig") as f:
            self.config: Dict[str, Any] = json.load(f)

    def execute_update(
        self,
        pptx_stream: BinaryIO,
        records: List[HeadcountRecord],
        reference_label: str = "Base Outubro 2026",
        original_filename: str = "apresentacao_original.pptx",
        custom_overrides: Optional[List[Dict[str, Any]]] = None,
    ) -> tuple[bytes, ExecutionReport]:
        prs = Presentation(pptx_stream)
        slides_before = len(prs.slides)

        total_company = sum(r.headcount for r in records)
        seg_totals: Dict[str, int] = {}
        area_totals: Dict[tuple[str, str], int] = {}
        unit_totals: Dict[tuple[str, str, str], int] = {}

        for r in records:
            seg_totals[r.segmento_des] = seg_totals.get(r.segmento_des, 0) + r.headcount
            area_totals[(r.segmento_des, r.area_des)] = (
                area_totals.get((r.segmento_des, r.area_des), 0) + r.headcount
            )
            unit_totals[(r.segmento_des, r.area_des, r.unidade_des)] = r.headcount

        report = ExecutionReport(
            original_filename=original_filename,
            output_filename=f"Atualizado_{original_filename}",
            slides_before=slides_before,
            slides_after=slides_before,
            total_company_headcount=total_company,
            reference_label=reference_label,
        )

        # 1. Datas Base
        for d in self.config.get("date_shapes", []):
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[d["slide_index"]], d["shape_path"])
            if sh:
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, reference_label)
                report.changes.append(
                    ChangeLogEntry(
                        d["slide_index"] + 1, str(d["shape_path"]), sh.name,
                        TargetType.REFERENCE_DATE, old_v, new_v, UpdateStatus.UPDATED_NUMBER
                    )
                )

        # Limpa a nota antiga de maio/junho no Slide 7 (índice 6, Shape 38), se existir
        obs_sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[6], "38")
        if obs_sh and getattr(obs_sh, "has_text_frame", False):
            old_obs = obs_sh.text_frame.text.strip()
            if "quadro de maio" in old_obs.lower():
                PptxSurgicalEngine.update_numeric_badge(obs_sh, "")
                report.changes.append(
                    ChangeLogEntry(
                        7, "38", obs_sh.name, TargetType.REFERENCE_DATE,
                        old_obs, "(Nota antiga de maio removida)", UpdateStatus.UPDATED_NAME_AND_NUMBER
                    )
                )

        # 2. Total Geral CNI
        for c in self.config.get("company_total_shapes", []):
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[c["slide_index"]], c["shape_path"])
            if sh:
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, total_company)
                report.changes.append(
                    ChangeLogEntry(
                        c["slide_index"] + 1, str(c["shape_path"]), sh.name,
                        TargetType.HEADCOUNT_COMPANY_TOTAL, old_v, new_v, UpdateStatus.UPDATED_NUMBER
                    )
                )

        # 3. Totais de Diretoria / Segmento
        for s in self.config.get("segment_total_shapes", []):
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[s["slide_index"]], s["shape_path"])
            if sh and s["segmento"] in seg_totals:
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, seg_totals[s["segmento"]])
                st = UpdateStatus.UPDATED_NUMBER if old_v != new_v else UpdateStatus.UNCHANGED
                report.changes.append(
                    ChangeLogEntry(
                        s["slide_index"] + 1, str(s["shape_path"]), s["segmento"],
                        TargetType.HEADCOUNT_SEGMENT_TOTAL, old_v, new_v, st
                    )
                )

        # 4. Subtotais de Coluna
        for a in self.config.get("area_subtotal_shapes", []):
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[a["slide_index"]], a["shape_path"])
            if sh:
                val = area_totals.get((a["segmento"], a["area"]), 0)
                if a["slide_index"] == 8 and str(a["shape_path"]) == "44":
                    val = area_totals.get(("SENAI", "Dados e Gestão"), 0) + area_totals.get(("SENAI", "Negócios Institucionais"), 0)
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, val)
                st = UpdateStatus.UPDATED_NUMBER if old_v != new_v else UpdateStatus.UNCHANGED
                report.changes.append(
                    ChangeLogEntry(
                        a["slide_index"] + 1, str(a["shape_path"]), a["area"],
                        TargetType.HEADCOUNT_AREA_SUBTOTAL, old_v, new_v, st
                    )
                )

        # 5. Caixas de Unidades Existentes
        for u in self.config.get("unit_boxes", []):
            sl = prs.slides[u["slide_index"]]
            name_sh = PptxSurgicalEngine.find_shape_by_path(sl, u["name_path"])
            badge_sh = PptxSurgicalEngine.find_shape_by_path(sl, u["badge_path"])
            if not name_sh or not badge_sh:
                continue

            val = unit_totals.get((u["segmento"], u["area"], u["unidade"]), 0)
            old_num, new_num = PptxSurgicalEngine.update_numeric_badge(badge_sh, val)
            old_name = name_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
            new_name = old_name

            if u.get("rename_to"):
                old_name, new_name = PptxSurgicalEngine.update_area_name_preserving_manager(name_sh, u["rename_to"])
                st = UpdateStatus.UPDATED_NAME_AND_NUMBER
            else:
                st = UpdateStatus.UPDATED_NUMBER if old_num != new_num else UpdateStatus.UNCHANGED

            report.changes.append(
                ChangeLogEntry(
                    u["slide_index"] + 1, f"{u['name_path']} | {u['badge_path']}", u["unidade"],
                    TargetType.HEADCOUNT_UNIT, f"{old_name} [{old_num}]", f"{new_name} [{new_num}]", st
                )
            )

        # 6. Novas Caixas Clonadas
        for cl in self.config.get("cloned_new_boxes", []):
            sl = prs.slides[cl["slide_index"]]
            val = unit_totals.get((cl["segmento"], cl["area"], cl["unidade"]), 0)
            cloned_info = PptxSurgicalEngine.clone_unit_group_in_column(
                slide=sl,
                template_group_id=int(cl["template_group_id"]),
                new_area_name=cl["display_name"],
                new_headcount=val,
                manager_placeholder=cl.get("manager_placeholder", "(Definir Gestor)"),
                left_override=cl.get("left_override"),
                top_override=cl.get("top_override"),
            )

            report.changes.append(
                ChangeLogEntry(
                    cl["slide_index"] + 1,
                    f"{cloned_info['name_path']} | {cloned_info['badge_path']}",
                    cl["display_name"],
                    TargetType.HEADCOUNT_UNIT,
                    "(Não existia no slide)",
                    f"{cl['display_name']} {cl.get('manager_placeholder', '(Definir Gestor)')} [{val}]",
                    UpdateStatus.CLONED_NEW_BOX,
                )
            )

        # 7. Caixas Extintas / Sem Contratos Ativos
        for ex in self.config.get("extinct_boxes", []):
            sl = prs.slides[ex["slide_index"]]
            name_sh = PptxSurgicalEngine.find_shape_by_path(sl, ex["name_path"])
            badge_sh = PptxSurgicalEngine.find_shape_by_path(sl, ex["badge_path"])
            if not name_sh or not badge_sh:
                continue
            old_num, _ = PptxSurgicalEngine.update_numeric_badge(badge_sh, 0)
            title = name_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
            report.changes.append(
                ChangeLogEntry(
                    ex["slide_index"] + 1, f"{ex['name_path']} | {ex['badge_path']}", title,
                    TargetType.HEADCOUNT_UNIT, f"{title} [{old_num}]", f"{title} [0]",
                    UpdateStatus.EXTINCT_IN_DATABRICKS, ex["reason"]
                )
            )

        # 8. APLICAÇÃO DE EDIÇÕES MANUAIS DA INTERFACE WEB (LIVE OVERRIDES)
        if custom_overrides:
            for ov in custom_overrides:
                s_idx = int(ov["slide_index"]) - 1  # Recebe 1-based da UI, converte para 0-based
                if not (0 <= s_idx < len(prs.slides)):
                    continue
                sl = prs.slides[s_idx]
                shape_key = str(ov.get("shape_id", "")).strip()

                if "|" in shape_key:
                    name_p, badge_p = [p.strip() for p in shape_key.split("|", 1)]
                    name_sh = PptxSurgicalEngine.find_shape_by_path(sl, name_p)
                    badge_sh = PptxSurgicalEngine.find_shape_by_path(sl, badge_p)
                    old_txt, new_txt = "", ""
                    if name_sh and (ov.get("new_area_name") or ov.get("new_manager")):
                        curr_area, _ = PptxSurgicalEngine.split_area_and_manager(name_sh)
                        target_area = ov.get("new_area_name") or curr_area
                        old_txt, new_txt = PptxSurgicalEngine.update_area_name_preserving_manager(
                            name_sh, target_area, new_manager_override=ov.get("new_manager")
                        )
                    old_num, new_num = "", ""
                    if badge_sh and ov.get("new_number") is not None and str(ov.get("new_number")) != "":
                        old_num, new_num = PptxSurgicalEngine.update_numeric_badge(badge_sh, ov["new_number"])

                    report.changes.insert(
                        0,
                        ChangeLogEntry(
                            s_idx + 1, shape_key, "Edição Manual Web",
                            TargetType.AREA_NAME,
                            f"{old_txt} [{old_num}]".strip(),
                            f"{new_txt} [{new_num}]".strip(),
                            UpdateStatus.UPDATED_NAME_AND_NUMBER,
                            "Editado manualmente pelo utilizador na interface Web"
                        )
                    )
                else:
                    sh = PptxSurgicalEngine.find_shape_by_path(sl, shape_key)
                    if sh and ov.get("new_number") is not None:
                        old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, ov["new_number"])
                        report.changes.insert(
                            0,
                            ChangeLogEntry(
                                s_idx + 1, shape_key, "Edição Manual Web",
                                TargetType.HEADCOUNT_UNIT, old_v, new_v,
                                UpdateStatus.UPDATED_NUMBER,
                                "Valor ajustado manualmente na interface Web"
                            )
                        )

        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        pptx_bytes = output_buffer.getvalue()

        prs_val = Presentation(io.BytesIO(pptx_bytes))
        report.slides_after = len(prs_val.slides)
        report.technical_validation_passed = (report.slides_before == report.slides_after)
        report.updated_count = sum(1 for c in report.changes if c.status != UpdateStatus.UNCHANGED)
        report.unchanged_count = sum(1 for c in report.changes if c.status == UpdateStatus.UNCHANGED)
        report.conflict_count = sum(1 for c in report.changes if c.status == UpdateStatus.EXTINCT_IN_DATABRICKS)

        return pptx_bytes, report