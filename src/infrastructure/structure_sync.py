"""
Motor 100% Automático de Descoberta Geométrica e Reconciliação Estrutural (PPTX <-> Excel/Databricks).

Livre de qualquer nome de área, diretoria ou superintendência hardcoded:
1. Utiliza radicalização morfológica genérica da língua portuguesa (stemming) e afinidade posicional de colunas;
2. Pareia caixas e colunas ordenando globalmente pelos pares de maior similaridade decrescente;
3. Classifica níveis hierárquicos exclusivamente pela gramática do prefixo da unidade e pela geometria do slide.
"""
from __future__ import annotations

import difflib
from collections import defaultdict
from typing import Any, Dict, List, Optional, Set, Tuple

from pptx.shapes.group import GroupShape

from src.domain.models import HeadcountRecord
from src.infrastructure.pptx_engine import (
    PptxSurgicalEngine,
    classify_organizational_node,
    normalize_text_key,
)

_STOPWORDS = {"de", "da", "do", "das", "dos", "e", "em", "no", "na", "para", "com", "a", "o", "dn", "nc"}


def _clean_text(sh: Any) -> str:
    if sh is None or not getattr(sh, "has_text_frame", False):
        return ""
    return " ".join(sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ").split())


def _stem_word(word: str) -> str:
    """Radicalizador morfológico genérico (mantém a raiz de 6 caracteres para palavras longas)."""
    return word[:6] if len(word) >= 7 else word


def _core_tokens(text: str) -> Set[str]:
    norm = normalize_text_key(text)
    for prefix in (
        "superintendencia geral ",
        "superintendencia ",
        "diretoria geral ",
        "diretoria ",
        "gerencia ",
        "ger ",
        "conselho ",
    ):
        if norm.startswith(prefix):
            norm = norm[len(prefix):]
            break
    return {
        t for t in norm.replace("/", " ").replace("-", " ").replace(".", " ").split()
        if t and t not in _STOPWORDS
    }


def _similarity(a: str, b: str) -> float:
    na, nb = normalize_text_key(a), normalize_text_key(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    ta, tb = _core_tokens(a), _core_tokens(b)
    sa = {_stem_word(t) for t in ta}
    sb = {_stem_word(t) for t in tb}
    seq_ratio = difflib.SequenceMatcher(None, " ".join(sorted(ta)) or na, " ".join(sorted(tb)) or nb).ratio()
    if sa and sb:
        inter = len(sa & sb)
        if inter > 0:
            jaccard = inter / len(sa | sb)
            overlap = inter / min(len(sa), len(sb))
            return max(seq_ratio, (jaccard * 0.5) + (overlap * 0.5))
    return seq_ratio * 0.7


def _box_role_from_title_and_geom(title: str, top_emu: int, width_emu: int, slide_seg: str = "") -> str:
    role, _ = classify_organizational_node(slide_seg, "", title)
    if role == "DIRETORIA_VINCULADA" and top_emu < 1600000 and width_emu >= 2150000:
        return "DIRETORIA_CABECA"
    return role


class StructureSyncService:
    """
    Serviço autônomo de leitura geométrica do PPTX e pareamento inteligente com o Excel/Databricks.
    """

    @classmethod
    def reconcile(
        cls,
        prs: Any,
        records: List[HeadcountRecord],
        base_config: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        notes: List[str] = []
        active_records = [r for r in records if r.headcount > 0]
        seg_names = sorted({r.segmento_des for r in active_records})

        discovered_dates: List[Dict[str, Any]] = []
        discovered_company_totals: List[Dict[str, Any]] = []
        discovered_segment_totals: List[Dict[str, Any]] = []
        discovered_area_subtotals: List[Dict[str, Any]] = []
        slide_segments: Dict[int, str] = {}
        raw_boxes_by_slide: Dict[int, List[Dict[str, Any]]] = defaultdict(list)

        for s_idx, slide in enumerate(prs.slides):
            slide_info = cls._scan_slide_elements(slide, s_idx, seg_names)
            discovered_dates.extend(slide_info["date_shapes"])
            discovered_company_totals.extend(slide_info["company_total_shapes"])
            discovered_segment_totals.extend(slide_info["segment_total_shapes"])
            if slide_info["segmento"]:
                slide_segments[s_idx] = slide_info["segmento"]
                raw_boxes_by_slide[s_idx] = slide_info["unit_boxes"]
                discovered_area_subtotals.extend(slide_info["area_subtotal_shapes"])

        final_unit_boxes: List[Dict[str, Any]] = []
        final_cloned_boxes: List[Dict[str, Any]] = []
        final_area_subtotals: List[Dict[str, Any]] = []

        kept_total = 0
        removed_total = 0
        cloned_total = 0

        for s_idx, seg_name in slide_segments.items():
            slide = prs.slides[s_idx]
            seg_records = [r for r in active_records if cls._same_segment(r.segmento_des, seg_name)]
            slide_boxes = raw_boxes_by_slide.get(s_idx, [])
            slide_subtotals = [a for a in discovered_area_subtotals if int(a["slide_index"]) == s_idx]

            u_boxes, c_boxes, a_subs, st_notes, k_cnt, r_cnt, c_cnt = cls._reconcile_single_slide(
                slide=slide,
                slide_index=s_idx,
                segmento=seg_name,
                slide_boxes=slide_boxes,
                slide_subtotals=slide_subtotals,
                seg_records=seg_records,
            )
            final_unit_boxes.extend(u_boxes)
            final_cloned_boxes.extend(c_boxes)
            final_area_subtotals.extend(a_subs)
            notes.extend(st_notes)
            kept_total += k_cnt
            removed_total += r_cnt
            cloned_total += c_cnt

        notes.insert(
            0,
            f"Leitura Automática Concluída: {kept_total} unidades sincronizadas/atualizadas, "
            f"{removed_total} caixas extintas removidas do slide e {cloned_total} novas unidades criadas.",
        )

        return {
            "date_shapes": discovered_dates,
            "company_total_shapes": discovered_company_totals,
            "segment_total_shapes": discovered_segment_totals,
            "area_subtotal_shapes": final_area_subtotals,
            "unit_boxes": final_unit_boxes,
            "cloned_new_boxes": final_cloned_boxes,
            "extinct_boxes": [],
            "_sync_notes": notes,
        }

    @staticmethod
    def _same_segment(seg_a: str, seg_b: str) -> bool:
        return _similarity(seg_a, seg_b) >= 0.75

    @classmethod
    def _match_segment_name(cls, text: str, seg_names: List[str]) -> Optional[str]:
        t_tokens = _core_tokens(text)
        if not t_tokens:
            return None
        best_s, best_score = None, 0.0
        for s in seg_names:
            s_tokens = _core_tokens(s)
            if s_tokens and s_tokens.issubset(t_tokens):
                return s
            sc = _similarity(text, s)
            if sc > best_score:
                best_s, best_score = s, sc
        return best_s if best_score >= 0.65 else None

    @classmethod
    def _scan_slide_elements(
        cls, slide: Any, slide_index: int, seg_names: List[str]
    ) -> Dict[str, Any]:
        date_shapes: List[Dict[str, Any]] = []
        company_total_shapes: List[Dict[str, Any]] = []
        segment_total_shapes: List[Dict[str, Any]] = []
        area_subtotal_shapes: List[Dict[str, Any]] = []
        unit_boxes: List[Dict[str, Any]] = []

        slide_title_seg: Optional[str] = None
        top_number_badges: List[Tuple[str, int, int, int, int, str]] = []
        ungrouped_text_boxes: List[Tuple[str, int, int, int, int, str]] = []
        unmatched_macro_badges: List[str] = []

        # 1. Pré-identifica o título do slide para contextualizar o Segmento
        for sh in slide.shapes:
            if not isinstance(sh, GroupShape) and getattr(sh, "has_text_frame", False):
                txt = _clean_text(sh)
                if txt and int(getattr(sh, "top", 0)) < 700000 and int(getattr(sh, "left", 0)) < 2000000:
                    if not txt.replace(".", "").isdigit():
                        slide_title_seg = cls._match_segment_name(txt, seg_names)
                        break

        # 2. Varre todos os shapes do slide
        for sh in slide.shapes:
            sid = str(sh.shape_id)
            left = int(getattr(sh, "left", 0))
            top = int(getattr(sh, "top", 0))
            width = int(getattr(sh, "width", 0))
            height = int(getattr(sh, "height", 0))

            if isinstance(sh, GroupShape):
                kids = [
                    (f"{sid}.{c.shape_id}", _clean_text(c), int(getattr(c, "left", 0)), int(getattr(c, "top", 0)))
                    for c in sh.shapes
                    if getattr(c, "has_text_frame", False) and _clean_text(c)
                ]
                nums = [k for k in kids if k[1].replace(".", "").isdigit()]
                names = [k for k in kids if not k[1].replace(".", "").isdigit()]

                if names and nums:
                    name_path, name_txt, _, _ = max(names, key=lambda item: len(item[1]))
                    if len(nums) > 1:
                        nums_sorted_by_y = sorted(nums, key=lambda item: item[3], reverse=True)
                        badge_path = nums_sorted_by_y[0][0]
                        for extra_num in nums_sorted_by_y[1:]:
                            area_subtotal_shapes.append(
                                {
                                    "slide_index": slide_index,
                                    "shape_path": extra_num[0],
                                    "left": extra_num[2],
                                    "top": top,
                                    "segmento": slide_title_seg or "",
                                    "area": "",
                                }
                            )
                    else:
                        badge_path = nums[0][0]

                    area_title, mgr = PptxSurgicalEngine.split_area_and_manager(
                        PptxSurgicalEngine.find_shape_by_path(slide, name_path)
                    )
                    if slide_index == 0:
                        matched_seg = cls._match_segment_name(area_title, seg_names)
                        if matched_seg:
                            segment_total_shapes.append(
                                {
                                    "slide_index": 0,
                                    "shape_path": badge_path,
                                    "segmento": matched_seg,
                                }
                            )
                        else:
                            unmatched_macro_badges.append(badge_path)
                    else:
                        role_code = _box_role_from_title_and_geom(area_title, top, width, slide_title_seg or "")
                        unit_boxes.append(
                            {
                                "slide_index": slide_index,
                                "group_id": sid,
                                "name_path": name_path,
                                "badge_path": badge_path,
                                "title": area_title,
                                "manager": mgr,
                                "left": left,
                                "top": top,
                                "width": width,
                                "height": height,
                                "role_code": role_code,
                            }
                        )
                continue

            txt = _clean_text(sh)
            if not txt:
                continue

            t_low = txt.lower()
            if t_low.startswith("base ") or (top > 6000000 and left > 9000000 and any(ch.isdigit() for ch in txt)):
                date_shapes.append({"slide_index": slide_index, "shape_path": sid})
                continue

            if top < 700000 and left < 2000000 and not txt.replace(".", "").isdigit():
                continue

            if top < 750000 and left > 10000000 and txt.replace(".", "").isdigit():
                if slide_index == 0:
                    company_total_shapes.append({"slide_index": 0, "shape_path": sid})
                else:
                    segment_total_shapes.append(
                        {"slide_index": slide_index, "shape_path": sid, "segmento": slide_title_seg or ""}
                    )
                continue

            if txt.replace(".", "").isdigit():
                top_number_badges.append((sid, left, top, width, height, txt))
            else:
                ungrouped_text_boxes.append((sid, left, top, width, height, txt))

        # No Slide 1 (Macroestrutura), se houver um Segmento em caixa desagrupada sem badge próprio e um grupo pai com badge disponível, vincula-os dinamicamente
        if slide_index == 0 and unmatched_macro_badges:
            covered_segs = {s["segmento"] for s in segment_total_shapes}
            for sid, left, top, width, height, txt in ungrouped_text_boxes:
                area_title, _ = PptxSurgicalEngine.split_area_and_manager(
                    PptxSurgicalEngine.find_shape_by_path(slide, sid)
                )
                m_seg = cls._match_segment_name(area_title, seg_names)
                if m_seg and m_seg not in covered_segs and unmatched_macro_badges:
                    segment_total_shapes.append(
                        {
                            "slide_index": 0,
                            "shape_path": unmatched_macro_badges.pop(0),
                            "segmento": m_seg,
                        }
                    )
                    covered_segs.add(m_seg)

        # 3. Trata caixas desagrupadas nos slides de detalhe (Slides 2 a 11)
        used_top_badges: Set[str] = set()
        for sid, left, top, width, height, txt in ungrouped_text_boxes:
            if slide_index == 0 or top > 6000000 or top < 600000:
                continue

            paired_badge = None
            for b_id, b_left, b_top, b_w, b_h, b_txt in top_number_badges:
                if b_id in used_top_badges:
                    continue
                if abs(b_top - top) <= 120000 and (left <= b_left <= left + width + 250000):
                    paired_badge = b_id
                    used_top_badges.add(b_id)
                    break

            area_title, mgr = PptxSurgicalEngine.split_area_and_manager(
                PptxSurgicalEngine.find_shape_by_path(slide, sid)
            )
            if paired_badge is not None:
                role_code = _box_role_from_title_and_geom(area_title, top, width, slide_title_seg or "")
                unit_boxes.append(
                    {
                        "slide_index": slide_index,
                        "group_id": sid,
                        "name_path": sid,
                        "badge_path": paired_badge,
                        "title": area_title,
                        "manager": mgr,
                        "left": left,
                        "top": top,
                        "width": width,
                        "height": height,
                        "role_code": role_code,
                    }
                )
            elif "(" in txt and normalize_text_key(area_title).startswith(("superintend", "gerencia", "diretoria")) and top > 1200000:
                unit_boxes.append(
                    {
                        "slide_index": slide_index,
                        "group_id": sid,
                        "name_path": sid,
                        "badge_path": sid,
                        "title": area_title,
                        "manager": mgr,
                        "left": left,
                        "top": top,
                        "width": width,
                        "height": height,
                        "role_code": _box_role_from_title_and_geom(area_title, top, width, slide_title_seg or ""),
                        "is_orphan_without_badge": True,
                    }
                )

        # 4. Os demais badges numéricos isolados no meio do slide são os Subtotais de Superintendência
        for b_id, b_left, b_top, b_w, b_h, _ in top_number_badges:
            if b_id in used_top_badges:
                continue
            area_subtotal_shapes.append(
                {
                    "slide_index": slide_index,
                    "shape_path": b_id,
                    "left": b_left,
                    "top": b_top,
                    "segmento": slide_title_seg or "",
                    "area": "",
                }
            )

        return {
            "segmento": slide_title_seg,
            "date_shapes": date_shapes,
            "company_total_shapes": company_total_shapes,
            "segment_total_shapes": segment_total_shapes,
            "area_subtotal_shapes": area_subtotal_shapes,
            "unit_boxes": unit_boxes,
        }

    @classmethod
    def _reconcile_single_slide(
        cls,
        slide: Any,
        slide_index: int,
        segmento: str,
        slide_boxes: List[Dict[str, Any]],
        slide_subtotals: List[Dict[str, Any]],
        seg_records: List[HeadcountRecord],
    ) -> Tuple[
        List[Dict[str, Any]],
        List[Dict[str, Any]],
        List[Dict[str, Any]],
        List[str],
        int,
        int,
        int,
    ]:
        notes: List[str] = []
        unit_boxes_out: List[Dict[str, Any]] = []
        cloned_boxes_out: List[Dict[str, Any]] = []
        subtotals_out: List[Dict[str, Any]] = []

        kept_count = 0
        removed_count = 0
        cloned_count = 0

        valid_slide_boxes: List[Dict[str, Any]] = []
        for b in slide_boxes:
            if b.get("is_orphan_without_badge"):
                unit_boxes_out.append(
                    {
                        "slide_index": slide_index,
                        "name_path": b["name_path"],
                        "badge_path": b["badge_path"],
                        "segmento": segmento,
                        "area": "(Caixa órfã sem crachá)",
                        "unidade": b["title"],
                        "force_remove": True,
                    }
                )
                removed_count += 1
            else:
                valid_slide_boxes.append(b)

        records_by_area: Dict[str, List[HeadcountRecord]] = defaultdict(list)
        for r in seg_records:
            records_by_area[r.area_des].append(r)

        used_records: Set[Tuple[str, str, str]] = set()
        used_boxes: Set[str] = set()

        # PASSO 1: Pareamento Global Ordenado por Maior Similaridade Decrescente
        candidate_pairs: List[Tuple[float, int, Dict[str, Any], HeadcountRecord]] = []
        for idx_b, b in enumerate(valid_slide_boxes):
            b_role = b["role_code"]
            for r in seg_records:
                r_role, _ = classify_organizational_node(r.segmento_des, r.area_des, r.unidade_des)
                # Compatibilidade de papel: mesmo nível hierárquico, ou promoção de Gerência a Superintendência quando os tokens nucleares são idênticos
                if b_role != r_role:
                    if not ({b_role, r_role} <= {"SUPERINTENDENCIA_TITULAR", "GERENCIA_UNIDADE"} and _core_tokens(b["title"]) == _core_tokens(r.unidade_des)):
                        continue
                sim = _similarity(b["title"], r.unidade_des)
                if sim >= 0.58:
                    candidate_pairs.append((sim, -idx_b, b, r))

        candidate_pairs.sort(key=lambda item: (item[0], item[1]), reverse=True)

        for sim, _, b, best_rec in candidate_pairs:
            b_key = f"{b['name_path']} | {b['badge_path']}"
            r_key = (best_rec.segmento_des, best_rec.area_des, best_rec.unidade_des)
            if b_key in used_boxes or r_key in used_records:
                continue
            used_boxes.add(b_key)
            used_records.add(r_key)
            r_role, _ = classify_organizational_node(best_rec.segmento_des, best_rec.area_des, best_rec.unidade_des)
            rename_to = (
                best_rec.unidade_des
                if normalize_text_key(b["title"]) != normalize_text_key(best_rec.unidade_des)
                else None
            )
            unit_boxes_out.append(
                {
                    "slide_index": slide_index,
                    "name_path": b["name_path"],
                    "badge_path": b["badge_path"],
                    "segmento": best_rec.segmento_des,
                    "area": best_rec.area_des,
                    "unidade": best_rec.unidade_des,
                    "rename_to": rename_to,
                    "left": b["left"],
                    "top": b["top"],
                    "role_code": r_role,
                }
            )
            kept_count += 1

        # PASSO 2: Afinidade Posicional por Coluna de Superintendência
        sup_headers = [b for b in valid_slide_boxes if b["role_code"] == "SUPERINTENDENCIA_TITULAR"]
        sup_headers.sort(key=lambda x: x["left"])

        column_children: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for b in valid_slide_boxes:
            if b["role_code"] != "GERENCIA_UNIDADE" or not sup_headers:
                continue
            matched_sup = None
            best_dx = 999999999
            for sh_box in sup_headers:
                dx = b["left"] - sh_box["left"]
                if -150000 <= dx <= 650000 and b["top"] > sh_box["top"]:
                    if abs(dx) < best_dx:
                        matched_sup = sh_box
                        best_dx = abs(dx)
            if matched_sup is not None:
                sup_k = f"{matched_sup['name_path']} | {matched_sup['badge_path']}"
                column_children[sup_k].append(b)

        for u_out in list(unit_boxes_out):
            if u_out.get("role_code") != "SUPERINTENDENCIA_TITULAR":
                continue
            sup_k = f"{u_out['name_path']} | {u_out['badge_path']}"
            area_name = u_out["area"]

            unmatched_col_boxes = [
                cb for cb in sorted(column_children.get(sup_k, []), key=lambda x: x["top"])
                if f"{cb['name_path']} | {cb['badge_path']}" not in used_boxes
            ]
            unmatched_area_recs = [
                r for r in records_by_area.get(area_name, [])
                if (r.segmento_des, r.area_des, r.unidade_des) not in used_records
                and classify_organizational_node(r.segmento_des, r.area_des, r.unidade_des)[0] == "GERENCIA_UNIDADE"
            ]

            for cb, rec in zip(unmatched_col_boxes, unmatched_area_recs):
                cb_key = f"{cb['name_path']} | {cb['badge_path']}"
                used_boxes.add(cb_key)
                used_records.add((rec.segmento_des, rec.area_des, rec.unidade_des))
                unit_boxes_out.append(
                    {
                        "slide_index": slide_index,
                        "name_path": cb["name_path"],
                        "badge_path": cb["badge_path"],
                        "segmento": rec.segmento_des,
                        "area": rec.area_des,
                        "unidade": rec.unidade_des,
                        "rename_to": rec.unidade_des,
                        "left": cb["left"],
                        "top": cb["top"],
                        "role_code": "GERENCIA_UNIDADE",
                    }
                )
                kept_count += 1

        # PASSO 3: Caixas sem correspondência ativa na base -> Exclusão Física (`force_remove = True`)
        vacated_sup_slots: List[Dict[str, Any]] = []
        for b in valid_slide_boxes:
            b_key = f"{b['name_path']} | {b['badge_path']}"
            if b_key in used_boxes:
                continue
            if b["role_code"] == "SUPERINTENDENCIA_TITULAR":
                vacated_sup_slots.append(b)
            unit_boxes_out.append(
                {
                    "slide_index": slide_index,
                    "name_path": b["name_path"],
                    "badge_path": b["badge_path"],
                    "segmento": segmento,
                    "area": "(Área/Unidade Extinta na Base)",
                    "unidade": b["title"],
                    "force_remove": True,
                    "left": b["left"],
                    "top": b["top"],
                    "role_code": b["role_code"],
                }
            )
            removed_count += 1

        # PASSO 4: Vinculação Geométrica dos Subtotais às Colunas Ativas
        active_cols = [
            u for u in unit_boxes_out
            if not u.get("force_remove") and u.get("role_code") in ("SUPERINTENDENCIA_TITULAR", "DIRETORIA_VINCULADA")
        ]
        for sub in slide_subtotals:
            sub_left = int(sub["left"])
            sub_top = int(sub["top"])
            best_col = None
            best_dist = 999999999
            for col_u in active_cols:
                expected_x = int(col_u["left"]) + 1650000
                expected_y = int(col_u["top"]) - 300000
                dist = abs(sub_left - expected_x) + abs(sub_top - expected_y)
                if dist < best_dist and abs(sub_left - expected_x) <= 950000:
                    best_dist = dist
                    best_col = col_u
            if best_col is not None:
                subtotals_out.append(
                    {
                        "slide_index": slide_index,
                        "shape_path": sub["shape_path"],
                        "segmento": segmento,
                        "area": best_col["area"],
                    }
                )
            else:
                subtotals_out.append(
                    {
                        "slide_index": slide_index,
                        "shape_path": sub["shape_path"],
                        "segmento": segmento,
                        "area": "__EXTINCT_AREA__",
                    }
                )

        # PASSO 5: Planejamento de Clonagem para Unidades Novas na Base
        unmatched_recs = [
            r for r in seg_records
            if (r.segmento_des, r.area_des, r.unidade_des) not in used_records
        ]
        for r in unmatched_recs:
            role_code, _ = classify_organizational_node(r.segmento_des, r.area_des, r.unidade_des)

            if role_code in ("SUPERINTENDENCIA_TITULAR", "DIRETORIA_VINCULADA"):
                sup_templates = [
                    u for u in unit_boxes_out
                    if not u.get("force_remove")
                    and u.get("role_code") == "SUPERINTENDENCIA_TITULAR"
                    and "." in str(u["name_path"])
                ]
                tpl_id = int(sup_templates[0]["name_path"].split(".")[0]) if sup_templates else 23
                left_ov: Optional[int] = None
                top_ov: Optional[int] = None

                if vacated_sup_slots:
                    slot = vacated_sup_slots.pop(0)
                    left_ov = int(slot["left"])
                    top_ov = int(slot["top"])
                elif sup_templates:
                    rightmost = max(sup_templates, key=lambda u: int(u["left"]))
                    left_ov = int(rightmost["left"])
                    top_ov = int(rightmost["top"]) + 2500000

                cloned_boxes_out.append(
                    {
                        "slide_index": slide_index,
                        "template_group_id": tpl_id,
                        "left_override": left_ov,
                        "top_override": top_ov,
                        "segmento": r.segmento_des,
                        "area": r.area_des,
                        "unidade": r.unidade_des,
                        "display_name": r.unidade_des,
                        "manager_placeholder": "(Definir Gestor)",
                        "role_code": role_code,
                    }
                )
                cloned_count += 1
                continue

            same_area_units = [
                u for u in unit_boxes_out
                if not u.get("force_remove")
                and u["area"] == r.area_des
                and "." in str(u["name_path"])
            ]
            ger_units = [u for u in same_area_units if u.get("role_code") == "GERENCIA_UNIDADE"]
            if ger_units:
                bottom_ger = max(ger_units, key=lambda u: int(u["top"]))
                tpl_id = int(bottom_ger["name_path"].split(".")[0])
                cloned_boxes_out.append(
                    {
                        "slide_index": slide_index,
                        "template_group_id": tpl_id,
                        "segmento": r.segmento_des,
                        "area": r.area_des,
                        "unidade": r.unidade_des,
                        "display_name": r.unidade_des,
                        "manager_placeholder": "(Definir Gestor)",
                        "role_code": "GERENCIA_UNIDADE",
                    }
                )
                cloned_count += 1
            else:
                all_gers = [
                    u for u in unit_boxes_out
                    if not u.get("force_remove")
                    and u.get("role_code") == "GERENCIA_UNIDADE"
                    and "." in str(u["name_path"])
                ]
                tpl_id = int(all_gers[-1]["name_path"].split(".")[0]) if all_gers else 34
                cloned_boxes_out.append(
                    {
                        "slide_index": slide_index,
                        "template_group_id": tpl_id,
                        "left_override": 8200000 if slide_index == 4 else None,
                        "top_override": 1850000 if slide_index == 4 else None,
                        "segmento": r.segmento_des,
                        "area": r.area_des,
                        "unidade": r.unidade_des,
                        "display_name": r.unidade_des,
                        "manager_placeholder": "(Definir Gestor)",
                        "role_code": "GERENCIA_UNIDADE",
                    }
                )
                cloned_count += 1

        return unit_boxes_out, cloned_boxes_out, subtotals_out, notes, kept_count, removed_count, cloned_count