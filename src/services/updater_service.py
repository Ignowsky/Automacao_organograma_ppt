import io
import os
import json
from typing import List, Dict, Any, BinaryIO, Optional
from pptx import Presentation
from pptx.slide import Slide
from src.domain.models import (
    HeadcountRecord,
    ChangeLogEntry,
    ExecutionReport,
    UpdateStatus,
    TargetType,
)
from src.infrastructure.pptx_engine import PptxSurgicalEngine
from src.infrastructure.databricks_repository import parse_reference_date_info


class PresentationUpdaterService:
    """
    Orquestrador central das Fases 2 a 9 com suporte a:
    - Edição Manual (UPDATE), Adição de Novas Caixas (ADD) e Remoção de Caixas (REMOVE);
    - Recálculo Matemático em Cascata (Unidade -> Subtotal Superintendência -> Total Diretoria -> Total CNI);
    - Extração de Pré-Visualização 16:9 Interativa dos Slides (slides_preview).
    """

    def __init__(self, mapping_json_path: str = "config/slide_mapping.json") -> None:
        if not os.path.exists(mapping_json_path):
            raise FileNotFoundError(f"Arquivo de mapeamento não encontrado: {mapping_json_path}")
        if os.path.getsize(mapping_json_path) == 0:
            raise ValueError(f"O arquivo '{mapping_json_path}' está vazio (0 bytes).")

        with open(mapping_json_path, "r", encoding="utf-8-sig") as f:
            self.config: Dict[str, Any] = json.load(f)

    @staticmethod
    def _get_shape_left_emu(slide: Slide, shape_path: str) -> Optional[int]:
        """Obtém a coordenada horizontal (left) da shape ou de seu grupo pai no slide."""
        top_id = str(shape_path).split("|")[0].strip().split(".")[0].strip()
        top_sh = PptxSurgicalEngine.find_shape_by_path(slide, top_id)
        if top_sh is not None and hasattr(top_sh, "left"):
            return int(top_sh.left)
        sh = PptxSurgicalEngine.find_shape_by_path(slide, shape_path)
        if sh is not None and hasattr(sh, "left"):
            return int(sh.left)
        return None

    def _infer_segment_and_area_by_column(
        self, slide: Slide, slide_index: int, shape_key: str
    ) -> tuple[Optional[str], Optional[str]]:
        """
        Identifica a qual Diretoria (segmento) e Superintendência (area) uma caixa pertence
        comparando o alinhamento horizontal (left) com os subtotais e caixas da mesma coluna.
        """
        probe_path = shape_key.split("|")[0].strip() if "|" in shape_key else shape_key
        target_left = self._get_shape_left_emu(slide, probe_path)

        candidates: List[tuple[int, str, str]] = []
        for a in self.config.get("area_subtotal_shapes", []):
            if int(a["slide_index"]) == slide_index:
                a_left = self._get_shape_left_emu(slide, str(a["shape_path"]))
                if a_left is not None and target_left is not None:
                    candidates.append((abs(a_left - target_left), a["segmento"], a["area"]))

        for u in self.config.get("unit_boxes", []):
            if int(u["slide_index"]) == slide_index:
                u_left = self._get_shape_left_emu(slide, str(u["name_path"]))
                if u_left is not None and target_left is not None:
                    candidates.append((abs(u_left - target_left), u["segmento"], u["area"]))

        if candidates:
            candidates.sort(key=lambda item: item[0])
            return candidates[0][1], candidates[0][2]

        for s in self.config.get("segment_total_shapes", []):
            if int(s["slide_index"]) == slide_index:
                return s["segmento"], None

        return None, None

    def execute_update(
        self,
        pptx_stream: BinaryIO,
        records: List[HeadcountRecord],
        reference_label: str = "",
        original_filename: str = "apresentacao_original.pptx",
        custom_overrides: Optional[List[Dict[str, Any]]] = None,
        auto_layout: bool = True,
    ) -> tuple[bytes, ExecutionReport]:
        prs = Presentation(pptx_stream)
        slides_before = len(prs.slides)

        if not reference_label or not reference_label.strip():
            sample_ref = records[0].ano_mes_referencia if records else None
            reference_label = parse_reference_date_info(sample_ref)["label"]

        # -------------------------------------------------------------------------
        # FASE A: TOTAIS BASE DA PLANILHA / DATABRICKS
        # -------------------------------------------------------------------------
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

        base_total_company = total_company
        base_seg_totals = dict(seg_totals)
        base_area_totals = dict(area_totals)

        # -------------------------------------------------------------------------
        # FASE B: PRÉ-PROCESSAMENTO DE OVERRIDES (UPDATE / ADD / REMOVE) + CASCATA
        # -------------------------------------------------------------------------
        unit_lookup: Dict[tuple[int, str], Dict[str, Any]] = {}
        for u in self.config.get("unit_boxes", []):
            s_i = int(u["slide_index"])
            full_k = f"{u['name_path']} | {u['badge_path']}"
            unit_lookup[(s_i, full_k)] = u
            unit_lookup[(s_i, str(u["badge_path"]).strip())] = u
            unit_lookup[(s_i, str(u["name_path"]).strip())] = u

        area_sub_lookup: Dict[tuple[int, str], Dict[str, Any]] = {}
        for a in self.config.get("area_subtotal_shapes", []):
            area_sub_lookup[(int(a["slide_index"]), str(a["shape_path"]).strip())] = a

        seg_tot_lookup: Dict[tuple[int, str], Dict[str, Any]] = {}
        for s in self.config.get("segment_total_shapes", []):
            seg_tot_lookup[(int(s["slide_index"]), str(s["shape_path"]).strip())] = s

        extinct_lookup: Dict[tuple[int, str], Dict[str, Any]] = {}
        for ex in self.config.get("extinct_boxes", []):
            s_i = int(ex["slide_index"])
            full_k = f"{ex['name_path']} | {ex['badge_path']}"
            extinct_lookup[(s_i, full_k)] = ex
            extinct_lookup[(s_i, str(ex["badge_path"]).strip())] = ex

        unit_overrides_by_key: Dict[tuple[int, str], Dict[str, Any]] = {}
        removed_unit_keys: Dict[tuple[int, str], Dict[str, Any]] = {}
        cloned_overrides_by_unit: Dict[tuple[int, str], Dict[str, Any]] = {}
        direct_shape_overrides: Dict[tuple[int, str], Dict[str, Any]] = {}
        unmapped_overrides: List[Dict[str, Any]] = []
        custom_added_boxes: List[Dict[str, Any]] = []
        custom_removed_unmapped: List[Dict[str, Any]] = []

        cascaded_area_deltas: Dict[tuple[str, str], int] = {}
        cascaded_seg_deltas: Dict[str, int] = {}
        cascaded_company_delta: int = 0

        def _apply_cascade_delta(seg_name: Optional[str], area_name: Optional[str], delta_val: int) -> None:
            nonlocal total_company, cascaded_company_delta
            if delta_val == 0:
                return
            if seg_name and area_name:
                a_tuple = (seg_name, area_name)
                area_totals[a_tuple] = max(0, area_totals.get(a_tuple, 0) + delta_val)
                cascaded_area_deltas[a_tuple] = cascaded_area_deltas.get(a_tuple, 0) + delta_val
            if seg_name:
                seg_totals[seg_name] = max(0, seg_totals.get(seg_name, 0) + delta_val)
                cascaded_seg_deltas[seg_name] = cascaded_seg_deltas.get(seg_name, 0) + delta_val
            total_company = max(0, total_company + delta_val)
            cascaded_company_delta += delta_val

        if custom_overrides:
            for ov in custom_overrides:
                try:
                    s_idx = int(ov["slide_index"]) - 1
                except (KeyError, ValueError, TypeError):
                    continue
                if not (0 <= s_idx < len(prs.slides)):
                    continue

                action = str(ov.get("action") or "UPDATE").strip().upper()
                shape_key = str(ov.get("shape_id", "")).strip()
                raw_num = ov.get("new_number")
                has_num = raw_num is not None and str(raw_num).strip() != ""
                new_num_int: Optional[int] = None
                if has_num:
                    try:
                        new_num_int = int(round(float(str(raw_num).strip())))
                    except ValueError:
                        new_num_int = None

                # --- AÇÃO 1: ADICIONAR NOVA CAIXA NA COLUNA (ADD) ---
                if action == "ADD":
                    add_hc = new_num_int if new_num_int is not None else 0
                    u_ref = unit_lookup.get((s_idx, shape_key))
                    if u_ref is not None:
                        inf_seg, inf_area = u_ref["segmento"], u_ref["area"]
                    else:
                        inf_seg, inf_area = self._infer_segment_and_area_by_column(
                            prs.slides[s_idx], s_idx, shape_key
                        )
                    _apply_cascade_delta(inf_seg, inf_area, add_hc)
                    custom_added_boxes.append(
                        {
                            "slide_index": s_idx,
                            "template_shape_key": shape_key,
                            "new_area_name": (ov.get("new_area_name") or "Nova Unidade").strip(),
                            "new_manager": (ov.get("new_manager") or "(Definir Gestor)").strip(),
                            "new_headcount": add_hc,
                            "segmento": inf_seg or "",
                            "area": inf_area or "",
                        }
                    )
                    continue

                # --- AÇÃO 2: REMOVER CAIXA EXISTENTE DO SLIDE (REMOVE) ---
                if action == "REMOVE":
                    if (s_idx, shape_key) in unit_lookup:
                        u_rule = unit_lookup[(s_idx, shape_key)]
                        canonical_key = f"{u_rule['name_path']} | {u_rule['badge_path']}"
                        removed_unit_keys[(s_idx, canonical_key)] = ov
                        u_tuple = (u_rule["segmento"], u_rule["area"], u_rule["unidade"])
                        prev_hc = unit_totals.get(u_tuple, 0)
                        unit_totals[u_tuple] = 0
                        _apply_cascade_delta(u_rule["segmento"], u_rule["area"], -prev_hc)
                    elif (s_idx, shape_key) in extinct_lookup:
                        ex_rule = extinct_lookup[(s_idx, shape_key)]
                        canonical_ex = f"{ex_rule['name_path']} | {ex['badge_path']}"
                        removed_unit_keys[(s_idx, canonical_ex)] = ov
                    else:
                        custom_removed_unmapped.append(ov)
                        badge_probe = shape_key.split("|")[1].strip() if "|" in shape_key else shape_key
                        sh_probe = PptxSurgicalEngine.find_shape_by_path(prs.slides[s_idx], badge_probe)
                        prev_hc = 0
                        if sh_probe and getattr(sh_probe, "has_text_frame", False):
                            try:
                                prev_hc = int(float(sh_probe.text_frame.text.strip()))
                            except ValueError:
                                prev_hc = 0
                        inf_seg, inf_area = self._infer_segment_and_area_by_column(
                            prs.slides[s_idx], s_idx, shape_key
                        )
                        _apply_cascade_delta(inf_seg, inf_area, -prev_hc)
                    continue

                # --- AÇÃO 3: EDITAR CAIXA EXISTENTE (UPDATE) ---
                if (s_idx, shape_key) in unit_lookup:
                    u_rule = unit_lookup[(s_idx, shape_key)]
                    canonical_key = f"{u_rule['name_path']} | {u_rule['badge_path']}"
                    unit_overrides_by_key[(s_idx, canonical_key)] = ov

                    if new_num_int is not None:
                        u_tuple = (u_rule["segmento"], u_rule["area"], u_rule["unidade"])
                        prev_hc = unit_totals.get(u_tuple, 0)
                        delta = new_num_int - prev_hc
                        unit_totals[u_tuple] = new_num_int
                        _apply_cascade_delta(u_rule["segmento"], u_rule["area"], delta)
                    continue

                matched_clone = None
                for cl in self.config.get("cloned_new_boxes", []):
                    if int(cl["slide_index"]) == s_idx:
                        if (
                            shape_key == f"Clonado de {cl['template_group_id']}"
                            or (ov.get("new_area_name") and ov["new_area_name"].strip() == cl["display_name"])
                        ):
                            matched_clone = cl
                            break
                if matched_clone is not None:
                    cloned_overrides_by_unit[(s_idx, matched_clone["unidade"])] = ov
                    if new_num_int is not None:
                        u_tuple = (
                            matched_clone["segmento"],
                            matched_clone["area"],
                            matched_clone["unidade"],
                        )
                        prev_hc = unit_totals.get(u_tuple, 0)
                        delta = new_num_int - prev_hc
                        unit_totals[u_tuple] = new_num_int
                        _apply_cascade_delta(matched_clone["segmento"], matched_clone["area"], delta)
                    continue

                if (s_idx, shape_key) in area_sub_lookup:
                    a_rule = area_sub_lookup[(s_idx, shape_key)]
                    direct_shape_overrides[(s_idx, shape_key)] = ov
                    if new_num_int is not None:
                        a_tuple = (a_rule["segmento"], a_rule["area"])
                        if s_idx == 8 and shape_key == "44":
                            prev_sub = area_totals.get(("SENAI", "Dados e Gestão"), 0) + area_totals.get(
                                ("SENAI", "Negócios Institucionais"), 0
                            )
                        else:
                            prev_sub = area_totals.get(a_tuple, 0)
                        delta = new_num_int - prev_sub
                        _apply_cascade_delta(a_rule["segmento"], a_rule["area"], delta)
                    continue

                if (s_idx, shape_key) in seg_tot_lookup:
                    s_rule = seg_tot_lookup[(s_idx, shape_key)]
                    direct_shape_overrides[(s_idx, shape_key)] = ov
                    if new_num_int is not None:
                        seg_name = s_rule["segmento"]
                        prev_seg = seg_totals.get(seg_name, 0)
                        delta = new_num_int - prev_seg
                        _apply_cascade_delta(seg_name, None, delta)
                    continue

                if (s_idx, shape_key) in extinct_lookup:
                    ex_rule = extinct_lookup[(s_idx, shape_key)]
                    canonical_ex = f"{ex_rule['name_path']} | {ex_rule['badge_path']}"
                    direct_shape_overrides[(s_idx, canonical_ex)] = ov
                    prev_ex_hc = 0
                else:
                    unmapped_overrides.append(ov)
                    badge_probe = shape_key.split("|")[1].strip() if "|" in shape_key else shape_key
                    sh_probe = PptxSurgicalEngine.find_shape_by_path(prs.slides[s_idx], badge_probe)
                    prev_ex_hc = 0
                    if sh_probe and getattr(sh_probe, "has_text_frame", False):
                        try:
                            prev_ex_hc = int(float(sh_probe.text_frame.text.strip()))
                        except ValueError:
                            prev_ex_hc = 0

                if new_num_int is not None:
                    delta = new_num_int - prev_ex_hc
                    if delta != 0:
                        inf_seg, inf_area = self._infer_segment_and_area_by_column(
                            prs.slides[s_idx], s_idx, shape_key
                        )
                        _apply_cascade_delta(inf_seg, inf_area, delta)

        # -------------------------------------------------------------------------
        # FASE C: EXECUÇÃO DAS ATUALIZAÇÕES NO POWERPOINT E LOG DE AUDITORIA
        # -------------------------------------------------------------------------
        report = ExecutionReport(
            original_filename=original_filename,
            output_filename=f"Atualizado_{original_filename}",
            slides_before=slides_before,
            slides_after=slides_before,
            total_company_headcount=total_company,
            reference_label=reference_label,
        )

        manual_entries: List[ChangeLogEntry] = []
        cascade_entries: List[ChangeLogEntry] = []
        regular_entries: List[ChangeLogEntry] = []
        date_entries: List[ChangeLogEntry] = []

        # 1. Datas Base nos 11 Slides
        for d in self.config.get("date_shapes", []):
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[d["slide_index"]], d["shape_path"])
            if sh:
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, reference_label)
                st = UpdateStatus.UPDATED_NUMBER if old_v != new_v else UpdateStatus.UNCHANGED
                date_entries.append(
                    ChangeLogEntry(
                        d["slide_index"] + 1,
                        str(d["shape_path"]),
                        "Tag de Competência (Data Base)",
                        TargetType.REFERENCE_DATE,
                        old_v,
                        new_v,
                        st,
                        "Tag de competência nos slides",
                    )
                )

        if len(prs.slides) > 6:
            obs_sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[6], "38")
            if obs_sh and getattr(obs_sh, "has_text_frame", False):
                old_obs = obs_sh.text_frame.text.strip()
                if "quadro de maio" in old_obs.lower():
                    PptxSurgicalEngine.update_numeric_badge(obs_sh, "")
                    date_entries.append(
                        ChangeLogEntry(
                            7,
                            "38",
                            obs_sh.name,
                            TargetType.REFERENCE_DATE,
                            old_obs,
                            "(Nota antiga de maio removida)",
                            UpdateStatus.UPDATED_NAME_AND_NUMBER,
                            "Limpeza automática de observação antiga",
                        )
                    )

        # 2. Total Geral CNI
        for c in self.config.get("company_total_shapes", []):
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[c["slide_index"]], c["shape_path"])
            if sh:
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, total_company)
                if cascaded_company_delta != 0:
                    sign = "+" if cascaded_company_delta > 0 else ""
                    msg = (
                        f"🔄 Recalculado em cascata: Base ({base_total_company}) "
                        f"{sign}{cascaded_company_delta} (Ajustes Manuais) = {total_company}"
                    )
                    cascade_entries.append(
                        ChangeLogEntry(
                            c["slide_index"] + 1,
                            str(c["shape_path"]),
                            "Total Geral CNI (Macroestrutura)",
                            TargetType.HEADCOUNT_COMPANY_TOTAL,
                            f"Total Geral CNI [{old_v}]",
                            f"Total Geral CNI [{new_v}]",
                            UpdateStatus.UPDATED_NUMBER,
                            msg,
                        )
                    )
                else:
                    st = UpdateStatus.UPDATED_NUMBER if old_v != new_v else UpdateStatus.UNCHANGED
                    regular_entries.append(
                        ChangeLogEntry(
                            c["slide_index"] + 1,
                            str(c["shape_path"]),
                            "Total Geral CNI (Macroestrutura)",
                            TargetType.HEADCOUNT_COMPANY_TOTAL,
                            f"Total Geral CNI [{old_v}]",
                            f"Total Geral CNI [{new_v}]",
                            st,
                            "Total Geral da Confederação",
                        )
                    )

        # 3. Totais de Diretoria / Segmento
        for s in self.config.get("segment_total_shapes", []):
            s_idx = int(s["slide_index"])
            s_path = str(s["shape_path"]).strip()
            seg_name = s["segmento"]
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[s_idx], s_path)
            if sh and seg_name in seg_totals:
                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, seg_totals[seg_name])
                seg_delta = cascaded_seg_deltas.get(seg_name, 0)
                is_direct = (s_idx, s_path) in direct_shape_overrides

                if is_direct:
                    manual_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            s_path,
                            f"Total Diretoria · {seg_name}",
                            TargetType.HEADCOUNT_SEGMENT_TOTAL,
                            f"Total {seg_name} [{old_v}]",
                            f"Total {seg_name} [{new_v}]",
                            UpdateStatus.UPDATED_NUMBER,
                            f"✏️ Total da Diretoria editado manualmente ({seg_name})",
                        )
                    )
                elif seg_delta != 0:
                    sign = "+" if seg_delta > 0 else ""
                    base_s = base_seg_totals.get(seg_name, 0)
                    msg = (
                        f"🔄 Total Diretoria recalculado em cascata: Base ({base_s}) "
                        f"{sign}{seg_delta} = {seg_totals[seg_name]}"
                    )
                    cascade_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            s_path,
                            f"Total Diretoria · {seg_name}",
                            TargetType.HEADCOUNT_SEGMENT_TOTAL,
                            f"Total {seg_name} [{old_v}]",
                            f"Total {seg_name} [{new_v}]",
                            UpdateStatus.UPDATED_NUMBER,
                            msg,
                        )
                    )
                else:
                    st = UpdateStatus.UPDATED_NUMBER if old_v != new_v else UpdateStatus.UNCHANGED
                    regular_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            s_path,
                            f"Total Diretoria · {seg_name}",
                            TargetType.HEADCOUNT_SEGMENT_TOTAL,
                            f"Total {seg_name} [{old_v}]",
                            f"Total {seg_name} [{new_v}]",
                            st,
                            f"Total consolidado da Diretoria ({seg_name})",
                        )
                    )

        # 4. Subtotais de Superintendência (Colunas)
        for a in self.config.get("area_subtotal_shapes", []):
            s_idx = int(a["slide_index"])
            a_path = str(a["shape_path"]).strip()
            seg_name = a["segmento"]
            area_name = a["area"]
            sh = PptxSurgicalEngine.find_shape_by_path(prs.slides[s_idx], a_path)
            if sh:
                if s_idx == 8 and a_path == "44":
                    val = area_totals.get(("SENAI", "Dados e Gestão"), 0) + area_totals.get(
                        ("SENAI", "Negócios Institucionais"), 0
                    )
                    base_a = base_area_totals.get(("SENAI", "Dados e Gestão"), 0) + base_area_totals.get(
                        ("SENAI", "Negócios Institucionais"), 0
                    )
                    area_delta = cascaded_area_deltas.get(
                        ("SENAI", "Dados e Gestão"), 0
                    ) + cascaded_area_deltas.get(("SENAI", "Negócios Institucionais"), 0)
                    display_area = "Dados e Gestão + Negócios Institucionais"
                else:
                    val = area_totals.get((seg_name, area_name), 0)
                    base_a = base_area_totals.get((seg_name, area_name), 0)
                    area_delta = cascaded_area_deltas.get((seg_name, area_name), 0)
                    display_area = area_name

                old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, val)
                is_direct = (s_idx, a_path) in direct_shape_overrides

                if is_direct:
                    manual_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            a_path,
                            f"Subtotal Superintendência · {display_area}",
                            TargetType.HEADCOUNT_AREA_SUBTOTAL,
                            f"Subtotal {display_area} [{old_v}]",
                            f"Subtotal {display_area} [{new_v}]",
                            UpdateStatus.UPDATED_NUMBER,
                            f"✏️ Subtotal editado manualmente · {seg_name} › {display_area}",
                        )
                    )
                elif area_delta != 0:
                    sign = "+" if area_delta > 0 else ""
                    msg = (
                        f"🔄 Subtotal Superintendência recalculado em cascata: "
                        f"Base ({base_a}) {sign}{area_delta} = {val} · {seg_name}"
                    )
                    cascade_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            a_path,
                            f"Subtotal Superintendência · {display_area}",
                            TargetType.HEADCOUNT_AREA_SUBTOTAL,
                            f"Subtotal {display_area} [{old_v}]",
                            f"Subtotal {display_area} [{new_v}]",
                            UpdateStatus.UPDATED_NUMBER,
                            msg,
                        )
                    )
                else:
                    st = UpdateStatus.UPDATED_NUMBER if old_v != new_v else UpdateStatus.UNCHANGED
                    regular_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            a_path,
                            f"Subtotal Superintendência · {display_area}",
                            TargetType.HEADCOUNT_AREA_SUBTOTAL,
                            f"Subtotal {display_area} [{old_v}]",
                            f"Subtotal {display_area} [{new_v}]",
                            st,
                            f"Subtotal da Superintendência · {seg_name} › {display_area}",
                        )
                    )

        # 5. Caixas de Unidades Existentes (Gerências / Gabinetes / Titulares)
        for u in self.config.get("unit_boxes", []):
            s_idx = int(u["slide_index"])
            sl = prs.slides[s_idx]
            canonical_key = f"{u['name_path']} | {u['badge_path']}"

            # Se o utilizador pediu para REMOVER esta caixa pela interface web:
            if (s_idx, canonical_key) in removed_unit_keys:
                rem_info = PptxSurgicalEngine.remove_unit_group_in_column(
                    sl, canonical_key, compact_column=auto_layout
                )
                old_txt = rem_info["old_text"] if rem_info else u["unidade"]
                old_hc = rem_info["old_hc"] if rem_info else 0
                manual_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        canonical_key,
                        u["unidade"],
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_txt} [{old_hc}]",
                        "(Caixa Removida do Slide) [0]",
                        UpdateStatus.REMOVED_BOX,
                        f"🗑️ Caixa removida e coluna alinhada (+Recálculo em Cascata) · {u['segmento']} › {u['area']}",
                    )
                )
                continue

            name_sh = PptxSurgicalEngine.find_shape_by_path(sl, u["name_path"])
            badge_sh = PptxSurgicalEngine.find_shape_by_path(sl, u["badge_path"])
            if not name_sh or not badge_sh:
                continue

            ov = unit_overrides_by_key.get((s_idx, canonical_key))
            val = unit_totals.get((u["segmento"], u["area"], u["unidade"]), 0)
            old_num, new_num = PptxSurgicalEngine.update_numeric_badge(badge_sh, val)
            old_name = name_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
            new_name = old_name

            if ov is not None:
                if ov.get("new_area_name") or ov.get("new_manager"):
                    curr_area, _ = PptxSurgicalEngine.split_area_and_manager(name_sh)
                    target_area = ov.get("new_area_name") or u.get("rename_to") or curr_area
                    _, new_name = PptxSurgicalEngine.update_area_name_preserving_manager(
                        name_sh, target_area, new_manager_override=ov.get("new_manager")
                    )
                elif u.get("rename_to"):
                    _, new_name = PptxSurgicalEngine.update_area_name_preserving_manager(
                        name_sh, u["rename_to"]
                    )

                manual_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        canonical_key,
                        u["unidade"],
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_name} [{old_num}]",
                        f"{new_name} [{new_num}]",
                        UpdateStatus.UPDATED_NAME_AND_NUMBER,
                        f"✏️ Editado manualmente (+Recálculo em Cascata) · {u['segmento']} › {u['area']}",
                    )
                )
            else:
                if u.get("rename_to"):
                    old_name, new_name = PptxSurgicalEngine.update_area_name_preserving_manager(
                        name_sh, u["rename_to"]
                    )
                    st = UpdateStatus.UPDATED_NAME_AND_NUMBER
                else:
                    st = UpdateStatus.UPDATED_NUMBER if old_num != new_num else UpdateStatus.UNCHANGED

                regular_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        canonical_key,
                        u["unidade"],
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_name} [{old_num}]",
                        f"{new_name} [{new_num}]",
                        st,
                        f"Unidade · {u['segmento']} › {u['area']}",
                    )
                )

        # 6. Novas Caixas Clonadas do Mapeamento Padrão
        for cl in self.config.get("cloned_new_boxes", []):
            s_idx = int(cl["slide_index"])
            sl = prs.slides[s_idx]
            ov = cloned_overrides_by_unit.get((s_idx, cl["unidade"]))
            val = unit_totals.get((cl["segmento"], cl["area"], cl["unidade"]), 0)
            disp_name = (ov.get("new_area_name") if ov and ov.get("new_area_name") else None) or cl["display_name"]
            mgr_ph = (
                (ov.get("new_manager") if ov and ov.get("new_manager") else None)
                or cl.get("manager_placeholder", "(Definir Gestor)")
            )

            cloned_info = PptxSurgicalEngine.clone_unit_group_in_column(
                slide=sl,
                template_group_id=int(cl["template_group_id"]),
                new_area_name=disp_name,
                new_headcount=val,
                manager_placeholder=mgr_ph,
                left_override=cl.get("left_override"),
                top_override=cl.get("top_override"),
            )

            entry = ChangeLogEntry(
                s_idx + 1,
                f"{cloned_info['name_path']} | {cloned_info['badge_path']}",
                disp_name,
                TargetType.HEADCOUNT_UNIT,
                "(Não existia no slide)",
                f"{disp_name} {mgr_ph} [{val}]",
                UpdateStatus.CLONED_NEW_BOX,
                f"Nova caixa clonada · {cl['segmento']} › {cl['area']}",
            )
            if ov is not None:
                manual_entries.append(entry)
            else:
                regular_entries.append(entry)

        # 6B. Novas Caixas Adicionadas Manualmente pelo Utilizador na Interface Web (action == 'ADD')
        for add_req in custom_added_boxes:
            s_idx = int(add_req["slide_index"])
            sl = prs.slides[s_idx]
            tpl_key = str(add_req["template_shape_key"])
            tpl_group_str = tpl_key.split("|")[0].strip().split(".")[0].strip()

            if tpl_group_str.isdigit():
                try:
                    cloned_info = PptxSurgicalEngine.clone_unit_group_in_column(
                        slide=sl,
                        template_group_id=int(tpl_group_str),
                        new_area_name=add_req["new_area_name"],
                        new_headcount=add_req["new_headcount"],
                        manager_placeholder=add_req["new_manager"],
                    )
                    manual_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            f"{cloned_info['name_path']} | {cloned_info['badge_path']}",
                            add_req["new_area_name"],
                            TargetType.HEADCOUNT_UNIT,
                            "(Nova Caixa Adicionada na Web)",
                            f"{add_req['new_area_name']} {add_req['new_manager']} [{add_req['new_headcount']}]",
                            UpdateStatus.CLONED_NEW_BOX,
                            f"➕ Nova caixa criada pelo utilizador (+Recálculo em Cascata) · {add_req['segmento']} › {add_req['area']}",
                        )
                    )
                except Exception:
                    pass

        # 7. Caixas Extintas / Sem Contratos Ativos
        for ex in self.config.get("extinct_boxes", []):
            s_idx = int(ex["slide_index"])
            sl = prs.slides[s_idx]
            canonical_ex = f"{ex['name_path']} | {ex['badge_path']}"

            if (s_idx, canonical_ex) in removed_unit_keys:
                rem_info = PptxSurgicalEngine.remove_unit_group_in_column(
                    sl, canonical_ex, compact_column=auto_layout
                )
                old_txt = rem_info["old_text"] if rem_info else "Unidade Extinta"
                manual_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        canonical_ex,
                        old_txt,
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_txt} [0]",
                        "(Caixa Removida do Slide) [0]",
                        UpdateStatus.REMOVED_BOX,
                        "🗑️ Caixa sem contratos removida do slide e coluna compactada",
                    )
                )
                continue

            name_sh = PptxSurgicalEngine.find_shape_by_path(sl, ex["name_path"])
            badge_sh = PptxSurgicalEngine.find_shape_by_path(sl, ex["badge_path"])
            if not name_sh or not badge_sh:
                continue

            ov = direct_shape_overrides.get((s_idx, canonical_ex))
            target_num = 0
            if ov is not None and ov.get("new_number") is not None and str(ov.get("new_number")).strip() != "":
                try:
                    target_num = int(round(float(str(ov["new_number"]).strip())))
                except ValueError:
                    target_num = 0

            old_num, new_num = PptxSurgicalEngine.update_numeric_badge(badge_sh, target_num)
            old_title = name_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
            new_title = old_title

            if ov is not None and (ov.get("new_area_name") or ov.get("new_manager")):
                curr_area, _ = PptxSurgicalEngine.split_area_and_manager(name_sh)
                target_area = ov.get("new_area_name") or curr_area
                _, new_title = PptxSurgicalEngine.update_area_name_preserving_manager(
                    name_sh, target_area, new_manager_override=ov.get("new_manager")
                )

            if ov is not None:
                manual_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        canonical_ex,
                        new_title,
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_title} [{old_num}]",
                        f"{new_title} [{new_num}]",
                        UpdateStatus.UPDATED_NAME_AND_NUMBER,
                        "✏️ Caixa reativada/editada manualmente (+Recálculo em Cascata)",
                    )
                )
            else:
                regular_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        canonical_ex,
                        old_title,
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_title} [{old_num}]",
                        f"{old_title} [0]",
                        UpdateStatus.EXTINCT_IN_DATABRICKS,
                        ex["reason"],
                    )
                )

        # 8. Remoções ou Edições de Shapes Avulsas
        for rem_ov in custom_removed_unmapped:
            s_idx = int(rem_ov["slide_index"]) - 1
            sl = prs.slides[s_idx]
            shape_key = str(rem_ov.get("shape_id", "")).strip()
            rem_info = PptxSurgicalEngine.remove_unit_group_in_column(
                sl, shape_key, compact_column=auto_layout
            )
            if rem_info:
                manual_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        shape_key,
                        rem_info["old_text"] or "Caixa Removida",
                        TargetType.HEADCOUNT_UNIT,
                        f"{rem_info['old_text']} [{rem_info['old_hc']}]",
                        "(Caixa Removida do Slide) [0]",
                        UpdateStatus.REMOVED_BOX,
                        "🗑️ Caixa removida pelo utilizador (+Recálculo em Cascata)",
                    )
                )

        for ov in unmapped_overrides:
            s_idx = int(ov["slide_index"]) - 1
            sl = prs.slides[s_idx]
            shape_key = str(ov.get("shape_id", "")).strip()

            if "|" in shape_key:
                name_p, badge_p = [p.strip() for p in shape_key.split("|", 1)]
                name_sh = PptxSurgicalEngine.find_shape_by_path(sl, name_p)
                badge_sh = PptxSurgicalEngine.find_shape_by_path(sl, badge_p)
                old_txt = (
                    name_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
                    if name_sh and getattr(name_sh, "has_text_frame", False)
                    else ""
                )
                new_txt = old_txt
                if name_sh and (ov.get("new_area_name") or ov.get("new_manager")):
                    curr_area, _ = PptxSurgicalEngine.split_area_and_manager(name_sh)
                    target_area = ov.get("new_area_name") or curr_area
                    _, new_txt = PptxSurgicalEngine.update_area_name_preserving_manager(
                        name_sh, target_area, new_manager_override=ov.get("new_manager")
                    )
                old_num = (
                    badge_sh.text_frame.text.strip()
                    if badge_sh and getattr(badge_sh, "has_text_frame", False)
                    else ""
                )
                new_num = old_num
                if badge_sh and ov.get("new_number") is not None and str(ov.get("new_number")) != "":
                    old_num, new_num = PptxSurgicalEngine.update_numeric_badge(badge_sh, ov["new_number"])

                manual_entries.append(
                    ChangeLogEntry(
                        s_idx + 1,
                        shape_key,
                        new_txt or "Edição Manual Web",
                        TargetType.HEADCOUNT_UNIT,
                        f"{old_txt} [{old_num}]".strip(),
                        f"{new_txt} [{new_num}]".strip(),
                        UpdateStatus.UPDATED_NAME_AND_NUMBER,
                        "✏️ Editado manualmente (+Recálculo em Cascata)",
                    )
                )
            else:
                sh = PptxSurgicalEngine.find_shape_by_path(sl, shape_key)
                if sh and ov.get("new_number") is not None and str(ov.get("new_number")) != "":
                    old_v, new_v = PptxSurgicalEngine.update_numeric_badge(sh, ov["new_number"])
                    manual_entries.append(
                        ChangeLogEntry(
                            s_idx + 1,
                            shape_key,
                            "Edição Manual Web",
                            TargetType.HEADCOUNT_UNIT,
                            old_v,
                            new_v,
                            UpdateStatus.UPDATED_NUMBER,
                            "✏️ Valor ajustado manualmente (+Recálculo em Cascata)",
                        )
                    )

        report.changes = manual_entries + cascade_entries + regular_entries + date_entries

        if auto_layout:
            report.validation_notes.append(
                f"Auto-Layout verificado: Caixas sincronizadas com competência '{reference_label}'."
            )
        if cascaded_area_deltas or cascaded_seg_deltas:
            affected_areas = ", ".join(
                f"{area} ({'+' if d > 0 else ''}{d})" for (_, area), d in cascaded_area_deltas.items()
            )
            report.validation_notes.append(
                f"Recálculo em Cascata aplicado automaticamente nas Superintendências: {affected_areas} "
                f"➔ Total Geral CNI ajustado para {total_company}."
            )

        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        pptx_bytes = output_buffer.getvalue()

        prs_val = Presentation(io.BytesIO(pptx_bytes))
        report.slides_after = len(prs_val.slides)
        report.technical_validation_passed = report.slides_before == report.slides_after
        report.updated_count = sum(1 for c in report.changes if c.status != UpdateStatus.UNCHANGED)
        report.unchanged_count = sum(1 for c in report.changes if c.status == UpdateStatus.UNCHANGED)
        report.conflict_count = sum(
            1 for c in report.changes if c.status == UpdateStatus.EXTINCT_IN_DATABRICKS
        )
        # Extrai a Pré-Visualização 16:9 de todos os slides a partir do PPTX validado final
        report.slides_preview = PptxSurgicalEngine.extract_slides_preview(prs_val)

        return pptx_bytes, report