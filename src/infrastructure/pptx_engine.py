import copy
import unicodedata
from typing import Optional, Dict, Any, List, Tuple
from pptx.presentation import Presentation
from pptx.slide import Slide
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.oxml.ns import qn
from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def normalize_text_key(text: str) -> str:
    if not text:
        return ""
    clean = " ".join(str(text).strip().split()).lower()
    nfkd = unicodedata.normalize("NFKD", clean)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def classify_organizational_node(segmento: str, area: str, unidade: str) -> Tuple[str, int]:
    s_norm = normalize_text_key(segmento)
    a_norm = normalize_text_key(area)
    u_norm = normalize_text_key(unidade)

    if u_norm.startswith("conselho") or u_norm.startswith("presidencia"):
        return ("CONSELHO_SUPERIOR", 0)

    if (
        "diretoria geral" in u_norm
        or "superintendencia geral" in u_norm
        or u_norm == s_norm
        or (u_norm.startswith("diretoria") and a_norm == s_norm)
    ):
        return ("DIRETORIA_CABECA", 1)

    if u_norm.startswith("diretoria") or a_norm.startswith("diretoria"):
        return ("DIRETORIA_VINCULADA", 2)

    if u_norm.startswith("superintend") or "negocios institucionais" in a_norm:
        return ("SUPERINTENDENCIA_TITULAR", 3)

    return ("GERENCIA_UNIDADE", 4)


class PptxSurgicalEngine:
    DEFAULT_VERTICAL_GAP_EMU = 95000
    STANDARD_UNIT_HEIGHT_EMU = 629312
    COLUMN_ALIGNMENT_TOLERANCE_EMU = 180000

    @staticmethod
    def find_shape_by_path(slide: Slide, shape_path: str | int) -> Optional[BaseShape]:
        parts = [int(p) for p in str(shape_path).strip().split(".") if p.strip().isdigit()]
        if not parts:
            return None
        current_container = slide.shapes
        target_shape: Optional[BaseShape] = None

        for idx, target_id in enumerate(parts):
            found = None
            for shape in current_container:
                if shape.shape_id == target_id:
                    found = shape
                    break
            if found is None:
                return None
            target_shape = found
            if idx < len(parts) - 1:
                if hasattr(found, "shapes"):
                    current_container = found.shapes
                else:
                    return None
        return target_shape

    @staticmethod
    def split_area_and_manager(shape: Optional[BaseShape]) -> tuple[str, str]:
        if shape is None or not getattr(shape, "has_text_frame", False):
            return "", ""
        raw = shape.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
        if "(" in raw:
            idx = raw.find("(")
            return raw[:idx].strip(), raw[idx:].strip()
        return raw, ""

    @staticmethod
    def update_numeric_badge(shape: Optional[BaseShape], new_number: int | str) -> tuple[str, str]:
        if shape is None or not getattr(shape, "has_text_frame", False):
            return "", str(new_number)

        old_val = shape.text_frame.text.strip()
        new_val = str(new_number)

        if not shape.text_frame.paragraphs:
            shape.text_frame.text = new_val
            return old_val, new_val

        p0 = shape.text_frame.paragraphs[0]
        if p0.runs:
            p0.runs[0].text = new_val
            for extra_run in p0.runs[1:]:
                extra_run.text = ""
        else:
            p0.text = new_val

        return old_val, new_val

    @staticmethod
    def update_area_name_preserving_manager(
        shape: Optional[BaseShape],
        new_area_name: str,
        new_manager_override: Optional[str] = None,
    ) -> tuple[str, str]:
        if shape is None or not getattr(shape, "has_text_frame", False):
            return "", new_area_name

        full_old = shape.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
        if not shape.text_frame.paragraphs:
            return full_old, new_area_name

        p0 = shape.text_frame.paragraphs[0]
        p_el = p0._p
        children = list(p_el)
        r_elements = [el for el in children if el.tag == qn("a:r")]

        mgr_r_el = None
        for r_el in r_elements:
            t_el = r_el.find(qn("a:t"))
            if t_el is not None and t_el.text and "(" in t_el.text:
                mgr_r_el = r_el
                break

        if mgr_r_el is not None and r_elements and r_elements[0] != mgr_r_el:
            first_r = r_elements[0]
            t_first = first_r.find(qn("a:t"))
            if t_first is not None:
                t_first.text = new_area_name

            idx_first = children.index(first_r)
            idx_mgr = children.index(mgr_r_el)
            between = children[idx_first + 1 : idx_mgr]
            br_between = [el for el in between if el.tag == qn("a:br")]
            keep_br = br_between[-1] if br_between else None

            for el in between:
                if el is not keep_br:
                    p_el.remove(el)

            if new_manager_override is not None:
                mgr_formatted = new_manager_override.strip()
                if mgr_formatted and not mgr_formatted.startswith("("):
                    mgr_formatted = f"({mgr_formatted})"
                t_mgr = mgr_r_el.find(qn("a:t"))
                if t_mgr is not None:
                    t_mgr.text = mgr_formatted
                for extra_r in r_elements[r_elements.index(mgr_r_el) + 1 :]:
                    t_ex = extra_r.find(qn("a:t"))
                    if t_ex is not None:
                        t_ex.text = ""

        elif mgr_r_el is not None and r_elements and r_elements[0] == mgr_r_el:
            t_el = mgr_r_el.find(qn("a:t"))
            if t_el is not None and t_el.text:
                ppos = t_el.text.find("(")
                mgr_txt = new_manager_override if new_manager_override is not None else t_el.text[ppos:]
                if mgr_txt and not mgr_txt.startswith("("):
                    mgr_txt = f"({mgr_txt})"
                t_el.text = f"{new_area_name} {mgr_txt}".strip()
        else:
            if r_elements:
                t_first = r_elements[0].find(qn("a:t"))
                if t_first is not None:
                    if new_manager_override:
                        mgr_txt = new_manager_override.strip()
                        if not mgr_txt.startswith("("):
                            mgr_txt = f"({mgr_txt})"
                        t_first.text = f"{new_area_name} {mgr_txt}"
                    else:
                        t_first.text = new_area_name
                for extra_el in children[children.index(r_elements[0]) + 1 :]:
                    if extra_el.tag in (qn("a:r"), qn("a:br")):
                        p_el.remove(extra_el)

        full_new = shape.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
        return full_old, full_new

    @staticmethod
    def apply_dashed_border_xml(
        shape_or_group: Optional[BaseShape],
        hex_color: str = "164194",
        width_emu: int = 19050,
    ) -> None:
        if shape_or_group is None:
            return

        el = shape_or_group._element
        target_sp = el

        if el.tag.endswith("grpSp"):
            child_shapes = el.findall(f".//{{{P_NS}}}sp")
            if child_shapes:
                def _shape_width(sp_node) -> int:
                    ext = sp_node.find(f".//{{{A_NS}}}xfrm/{{{A_NS}}}ext")
                    return int(ext.get("cx", "0")) if ext is not None else 0
                target_sp = max(child_shapes, key=_shape_width)

        spPr = target_sp.find(f"{{{P_NS}}}spPr")
        if spPr is None:
            spPr = target_sp.find(f"{{{A_NS}}}spPr")
        if spPr is None:
            return

        ln = spPr.find(f"{{{A_NS}}}ln")
        if ln is None:
            ln = etree.SubElement(spPr, f"{{{A_NS}}}ln")

        ln.set("w", str(width_emu))

        no_fill = ln.find(f"{{{A_NS}}}noFill")
        if no_fill is not None:
            ln.remove(no_fill)

        solid_fill = ln.find(f"{{{A_NS}}}solidFill")
        if solid_fill is None:
            solid_fill = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
        else:
            solid_fill.clear()

        srgb = etree.SubElement(solid_fill, f"{{{A_NS}}}srgbClr")
        srgb.set("val", hex_color.replace("#", "").upper())

        prst_dash = ln.find(f"{{{A_NS}}}prstDash")
        if prst_dash is None:
            prst_dash = etree.SubElement(ln, f"{{{A_NS}}}prstDash")
        prst_dash.set("val", "dash")

    @classmethod
    def auto_discover_slide_metadata(cls, prs: Presentation, records: List[Any]) -> Dict[str, Any]:
        """
        Analisa a apresentação e a base de headcount em tempo de execução, 
        mapeando dinamicamente cada unidade ao seu shape correspondente no slide por similaridade textual.
        """
        discovered_units = []
        for s_idx, slide in enumerate(prs.slides):
            geom_map = cls._collect_slide_geometry(slide)
            for path, g in geom_map.items():
                sh = g["shape"]
                if not getattr(sh, "has_text_frame", False):
                    continue
                txt = sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
                if not txt:
                    continue
                # Se o texto corresponder a uma unidade da base, mapeia dinamicamente
                for r in records:
                    if normalize_text_key(r.unidade_des) in normalize_text_key(txt):
                        discovered_units.append({
                            "slide_index": s_idx,
                            "segmento": r.segmento_des,
                            "area": r.area_des,
                            "unidade": r.unidade_des,
                            "name_path": path,
                            "badge_path": path
                        })
                        break
        return {"unit_boxes": discovered_units}

    @classmethod
    def _find_any_unit_group_on_slide(cls, slide: Slide, preferred_id_str: str) -> Optional[GroupShape]:
        sh = cls.find_shape_by_path(slide, preferred_id_str)
        if isinstance(sh, GroupShape):
            text_children = [c for c in sh.shapes if getattr(c, "has_text_frame", False)]
            if len(text_children) >= 2:
                return sh

        parts = [p.strip() for p in preferred_id_str.split(".") if p.strip().isdigit()]
        if len(parts) >= 2:
            parent_path = ".".join(parts[:-1])
            parent_sh = cls.find_shape_by_path(slide, parent_path)
            if isinstance(parent_sh, GroupShape):
                return parent_sh

        for candidate in slide.shapes:
            if isinstance(candidate, GroupShape):
                t_ch = [c for c in candidate.shapes if getattr(c, "has_text_frame", False)]
                if len(t_ch) >= 2:
                    return candidate
        return None

    @classmethod
    def clone_unit_group_in_column(
        cls,
        slide: Slide,
        template_group_id: int | str,
        new_area_name: str,
        new_headcount: int,
        manager_placeholder: str = "(Definir Gestor)",
        left_override: Optional[int] = None,
        top_override: Optional[int] = None,
        role_code: str = "GERENCIA_UNIDADE",
    ) -> Dict[str, Any]:
        template_group = cls._find_any_unit_group_on_slide(slide, str(template_group_id))
        if template_group is None:
            raise ValueError(f"Grupo base {template_group_id} não encontrado para clonagem.")

        ref_sh = cls.find_shape_by_path(slide, str(template_group_id).split(".")[0])
        col_left = int(getattr(ref_sh, "left", template_group.left))

        new_el = copy.deepcopy(template_group.element)
        max_id = max((s.shape_id for s in slide.shapes), default=200)
        new_group_id = max_id + 100
        new_el.nvGrpSpPr.cNvPr.id = new_group_id
        new_el.nvGrpSpPr.cNvPr.name = f"Agrupar_Auto_{new_group_id}"

        slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
        cloned_group: GroupShape = slide.shapes[-1]

        target_left = left_override if left_override is not None else col_left
        if left_override is None and role_code in ("SUPERINTENDENCIA_TITULAR", "DIRETORIA_VINCULADA"):
            target_left = col_left

        if top_override is not None:
            target_top = top_override
        else:
            lowest_bottom = int(getattr(ref_sh, "top", template_group.top)) + int(
                getattr(ref_sh, "height", template_group.height)
            )
            for sh in slide.shapes:
                if sh.shape_id == new_group_id:
                    continue
                if isinstance(sh, GroupShape) and abs(int(sh.left) - target_left) <= cls.COLUMN_ALIGNMENT_TOLERANCE_EMU:
                    sh_bottom = int(sh.top) + int(sh.height)
                    if sh_bottom > lowest_bottom and int(sh.top) < 5600000:
                        lowest_bottom = sh_bottom
            target_top = lowest_bottom + cls.DEFAULT_VERTICAL_GAP_EMU

        cloned_group.left = target_left
        cloned_group.top = target_top
        cloned_group.width = template_group.width
        cloned_group.height = (
            cls.STANDARD_UNIT_HEIGHT_EMU if len(new_area_name) < 35 else template_group.height
        )

        text_shapes = [s for s in cloned_group.shapes if getattr(s, "has_text_frame", False)]
        if len(text_shapes) >= 2:
            name_shape, badge_shape = text_shapes[0], text_shapes[1]
            cls.update_area_name_preserving_manager(
                name_shape, new_area_name, new_manager_override=manager_placeholder
            )
            cls.update_numeric_badge(badge_shape, new_headcount)

        cls.apply_dashed_border_xml(cloned_group, hex_color="164194")

        return {
            "new_group_id": new_group_id,
            "name_path": f"{new_group_id}.{text_shapes[0].shape_id}" if len(text_shapes) >= 2 else str(new_group_id),
            "badge_path": f"{new_group_id}.{text_shapes[1].shape_id}" if len(text_shapes) >= 2 else str(new_group_id),
            "left": cloned_group.left,
            "top": cloned_group.top,
            "area_name": new_area_name,
            "headcount": new_headcount,
            "container_shape": cloned_group,
        }

    @classmethod
    def remove_shapes_with_backdrop(cls, slide: Slide, paths: List[str]) -> int:
        """Remove shapes (por caminho) e os shapes SEM texto que formam o fundo dessa caixa
        (retângulo/borda cujo centro cai dentro da área ocupada pelos shapes removidos)."""
        shapes = [cls.find_shape_by_path(slide, p) for p in paths]
        shapes = [sh for sh in shapes if sh is not None]
        if not shapes:
            return 0
        first_parts = str(paths[0]).strip().split(".")
        if len(first_parts) > 1:
            parent = cls.find_shape_by_path(slide, ".".join(first_parts[:-1]))
            container = parent.shapes if parent is not None and hasattr(parent, "shapes") else None
        else:
            container = slide.shapes

        left = min(int(sh.left) for sh in shapes)
        top = min(int(sh.top) for sh in shapes)
        right = max(int(sh.left) + int(sh.width) for sh in shapes)
        bottom = max(int(sh.top) + int(sh.height) for sh in shapes)

        removed = 0
        for sh in shapes:
            el = sh.element
            if el.getparent() is not None:
                el.getparent().remove(el)
                removed += 1

        if container is not None:
            tol = 20000
            for sib in list(container):
                if isinstance(sib, GroupShape):
                    continue
                if getattr(sib, "has_text_frame", False) and sib.text_frame.text.strip():
                    continue
                try:
                    cx = int(sib.left) + int(sib.width) / 2
                    cy = int(sib.top) + int(sib.height) / 2
                    area_sib = int(sib.width) * int(sib.height)
                except Exception:
                    continue
                area_union = max(1, (right - left) * (bottom - top))
                if (left - tol <= cx <= right + tol and top - tol <= cy <= bottom + tol
                        and area_sib <= area_union * 1.6):
                    el = sib.element
                    if el.getparent() is not None:
                        el.getparent().remove(el)
        return removed

    @classmethod
    def remove_unit_group_in_column(
        cls,
        slide: Slide,
        shape_key: str,
        compact_column: bool = True,
        protected_paths: Optional[set] = None,
    ) -> Optional[Dict[str, Any]]:
        first_part = shape_key.split("|")[0].strip()
        parts = [p.strip() for p in first_part.split(".") if p.strip().isdigit()]
        if not parts:
            return None

        old_text = ""
        old_hc = 0

        if "|" in shape_key:
            name_p, badge_p = [p.strip() for p in shape_key.split("|", 1)]
            name_sh = cls.find_shape_by_path(slide, name_p)
            badge_sh = cls.find_shape_by_path(slide, badge_p)
            if name_sh and getattr(name_sh, "has_text_frame", False):
                old_text = name_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
            if badge_sh and getattr(badge_sh, "has_text_frame", False):
                try:
                    old_hc = int(float(badge_sh.text_frame.text.strip()))
                except ValueError:
                    old_hc = 0

        # Grupo-contêiner: se o grupo de topo guarda OUTROS shapes mapeados (subtotais, outras
        # caixas...), remover o grupo inteiro apagaria dados alheios. Remove só esta caixa.
        if len(parts) == 2 and protected_paths and "|" in shape_key:
            own = {name_p, badge_p}
            others = {p for p in protected_paths if p.startswith(parts[0] + ".") and p not in own}
            if others:
                n = cls.remove_shapes_with_backdrop(slide, [name_p, badge_p])
                if n:
                    return {
                        "removed_root_id": parts[0],
                        "old_text": old_text,
                        "old_hc": old_hc,
                        "left": 0,
                        "top": 0,
                        "shifted_count": 0,
                        "partial_removal": True,
                    }
                return None

        if len(parts) >= 3:
            sub_group_path = ".".join(parts[:-1])
            target_node = cls.find_shape_by_path(slide, sub_group_path)
        elif len(parts) == 2:
            target_node = cls.find_shape_by_path(slide, parts[0])
        else:
            target_node = cls.find_shape_by_path(slide, parts[0])

        if target_node is None:
            return None

        removed_left = int(getattr(target_node, "left", 0))
        removed_top = int(getattr(target_node, "top", 0))
        removed_height = int(getattr(target_node, "height", cls.STANDARD_UNIT_HEIGHT_EMU))

        parent_xml = target_node.element.getparent()
        if parent_xml is not None:
            parent_xml.remove(target_node.element)

        if len(parts) == 1 and "|" in shape_key:
            badge_part = shape_key.split("|")[1].strip().split(".")[0].strip()
            badge_node = cls.find_shape_by_path(slide, badge_part)
            if badge_node is not None and badge_node.element.getparent() is not None:
                badge_node.element.getparent().remove(badge_node.element)

        shifted_count = 0
        if compact_column and removed_height > 0 and parent_xml is not None:
            shift_delta = removed_height + cls.DEFAULT_VERTICAL_GAP_EMU
            for sh in slide.shapes:
                if isinstance(sh, GroupShape):
                    if (
                        abs(int(sh.left) - removed_left) <= cls.COLUMN_ALIGNMENT_TOLERANCE_EMU
                        and int(sh.top) > removed_top
                        and int(sh.top) < 5800000
                    ):
                        sh.top = max(0, int(sh.top) - shift_delta)
                        shifted_count += 1

        return {
            "removed_root_id": parts[0],
            "old_text": old_text,
            "old_hc": old_hc,
            "left": removed_left,
            "top": removed_top,
            "shifted_count": shifted_count,
        }

    @classmethod
    def _collect_slide_geometry(cls, slide: Slide) -> tuple[Dict[str, Dict[str, Any]], List[Dict[str, Any]]]:
        geom_map: Dict[str, Dict[str, Any]] = {}

        def _get_group_transform(grp_shape: GroupShape) -> tuple[float, float, float, float, float, float, float, float]:
            gx = float(getattr(grp_shape, "left", 0))
            gy = float(getattr(grp_shape, "top", 0))
            gw = float(getattr(grp_shape, "width", 1)) or 1.0
            gh = float(getattr(grp_shape, "height", 1)) or 1.0
            cx, cy, cw, ch = gx, gy, gw, gh
            try:
                xfrm = grp_shape.element.grpSpPr.xfrm
                if xfrm is not None:
                    ch_off = xfrm.find(qn("a:chOff"))
                    ch_ext = xfrm.find(qn("a:chExt"))
                    if ch_off is not None and ch_ext is not None:
                        cx = float(ch_off.get("x", 0))
                        cy = float(ch_off.get("y", 0))
                        cw = float(ch_ext.get("cx", 1)) or 1.0
                        ch = float(ch_ext.get("cy", 1)) or 1.0
            except Exception:
                pass
            return gx, gy, gw, gh, cx, cy, cw, ch

        def _walk(container, prefix: str = "", transform_fn=None) -> None:
            for sh in container:
                sid = str(sh.shape_id)
                path = f"{prefix}.{sid}" if prefix else sid

                raw_l = float(getattr(sh, "left", 0))
                raw_t = float(getattr(sh, "top", 0))
                raw_w = float(getattr(sh, "width", 0))
                raw_h = float(getattr(sh, "height", 0))

                if transform_fn is not None:
                    abs_l, abs_t, abs_w, abs_h = transform_fn(raw_l, raw_t, raw_w, raw_h)
                else:
                    abs_l, abs_t, abs_w, abs_h = raw_l, raw_t, raw_w, raw_h

                geom_map[path] = {
                    "shape": sh,
                    "path": path,
                    "left": abs_l,
                    "top": abs_t,
                    "width": abs_w,
                    "height": abs_h,
                    "is_group": isinstance(sh, GroupShape),
                }

                if isinstance(sh, GroupShape):
                    gx, gy, gw, gh, cx, cy, cw, ch = _get_group_transform(sh)

                    def _make_child_transform(p_fn, _gx=gx, _gy=gy, _gw=gw, _gh=gh, _cx=cx, _cy=cy, _cw=cw, _ch=ch):
                        def _t(l, t, w, h):
                            nl = _gx + ((l - _cx) * (_gw / _cw))
                            nt = _gy + ((t - _cy) * (_gh / _ch))
                            nw = w * (_gw / _cw)
                            nh = h * (_gh / _ch)
                            if p_fn is not None:
                                return p_fn(nl, nt, nw, nh)
                            return nl, nt, nw, nh

                        return _t

                    _walk(sh.shapes, path, _make_child_transform(transform_fn))

        _walk(slide.shapes)
        return geom_map

    @classmethod
    def extract_slides_preview(
        cls,
        prs: Presentation,
        mapping_config: Optional[Dict[str, Any]] = None,
    ) -> List[Dict[str, Any]]:
        slide_w = float(prs.slide_width or 12192000)
        slide_h = float(prs.slide_height or 6858000)
        cfg = mapping_config or {}
        slides_data: List[Dict[str, Any]] = []

        def _to_pct(l: float, t: float, w: float, h: float) -> Dict[str, float]:
            return {
                "left_pct": round(max(0.0, min(98.0, (l / slide_w) * 100.0)), 2),
                "top_pct": round(max(0.0, min(98.0, (t / slide_h) * 100.0)), 2),
                "width_pct": round(max(2.2, min(98.0, (w / slide_w) * 100.0)), 2),
                "height_pct": round(max(3.0, min(98.0, (h / slide_h) * 100.0)), 2),
            }

        for s_idx, slide in enumerate(prs.slides):
            geom_map = cls._collect_slide_geometry(slide)
            consumed_paths: set[str] = set()
            elements: List[Dict[str, Any]] = []

            def _mark_consumed(p: str) -> None:
                consumed_paths.add(p)
                parts = p.split(".")
                for i in range(1, len(parts)):
                    consumed_paths.add(".".join(parts[:i]))

            def _union_box(p1: str, p2: str) -> Optional[Dict[str, float]]:
                g1 = geom_map.get(p1)
                g2 = geom_map.get(p2)
                if not g1 and not g2:
                    return None
                if g1 and not g2:
                    return _to_pct(g1["left"], g1["top"], g1["width"], g1["height"])
                if g2 and not g1:
                    return _to_pct(g2["left"], g2["top"], g2["width"], g2["height"])
                min_l = min(g1["left"], g2["left"])
                min_t = min(g1["top"], g2["top"])
                max_r = max(g1["left"] + g1["width"], g2["left"] + g2["width"])
                max_b = max(g1["top"] + g1["height"], g2["top"] + g2["height"])
                return _to_pct(min_l, min_t, max_r - min_l, max_b - min_t)

            all_unit_rules = [
                {**u, "is_extinct": False} for u in cfg.get("unit_boxes", []) if int(u["slide_index"]) == s_idx
            ]

            for u in all_unit_rules:
                np_str = str(u["name_path"]).strip()
                bp_str = str(u["badge_path"]).strip()
                g_name = geom_map.get(np_str)
                g_badge = geom_map.get(bp_str)
                if not g_name or not g_badge:
                    continue

                _mark_consumed(np_str)
                _mark_consumed(bp_str)

                area_txt, mgr_txt = cls.split_area_and_manager(g_name["shape"])
                badge_txt = (
                    g_badge["shape"].text_frame.text.strip()
                    if getattr(g_badge["shape"], "has_text_frame", False)
                    else "0"
                )
                box_pct = _union_box(np_str, bp_str)
                if not box_pct:
                    continue

                if badge_txt == "0":
                    continue

                is_head = (
                    area_txt.lower().startswith("diretoria")
                    or area_txt.lower().startswith("superintend")
                    or area_txt.lower().startswith("conselho")
                )

                elements.append(
                    {
                        "kind": "UNIT_BOX",
                        "shape_key": f"{np_str} | {bp_str}",
                        "group_id": np_str.split(".")[0],
                        "segmento": u.get("segmento", ""),
                        "superintendencia": u.get("area", ""),
                        "area_name": area_txt,
                        "manager": mgr_txt,
                        "badge": badge_txt,
                        "is_head": is_head,
                        "is_extinct": False,
                        **box_pct,
                    }
                )

            for a in cfg.get("area_subtotal_shapes", []):
                if int(a["slide_index"]) != s_idx:
                    continue
                sp = str(a["shape_path"]).strip()
                g = geom_map.get(sp)
                if not g or not getattr(g["shape"], "has_text_frame", False):
                    continue
                _mark_consumed(sp)
                val_txt = g["shape"].text_frame.text.strip()
                pct = _to_pct(g["left"], g["top"], g["width"], g["height"])
                elements.append(
                    {
                        "kind": "AREA_SUBTOTAL",
                        "shape_key": sp,
                        "segmento": a.get("segmento", ""),
                        "superintendencia": a.get("area", ""),
                        "area_name": f"Subtotal · {a.get('area', '')}",
                        "badge": val_txt,
                        **pct,
                    }
                )

            tot_rules = [
                {**s, "label": f"Total {s.get('segmento', '')}", "is_company": False}
                for s in cfg.get("segment_total_shapes", [])
                if int(s["slide_index"]) == s_idx
            ] + [
                {**c, "label": "Total Geral CNI", "segmento": "CNI", "is_company": True}
                for c in cfg.get("company_total_shapes", [])
                if int(c["slide_index"]) == s_idx
            ]
            for t_rule in tot_rules:
                sp = str(t_rule["shape_path"]).strip()
                g = geom_map.get(sp)
                if not g or not getattr(g["shape"], "has_text_frame", False):
                    continue
                _mark_consumed(sp)
                val_txt = g["shape"].text_frame.text.strip()
                pct = _to_pct(g["left"], g["top"], g["width"], g["height"])
                elements.append(
                    {
                        "kind": "SLIDE_TOTAL_BADGE",
                        "shape_key": sp,
                        "area_name": t_rule["label"],
                        "badge": val_txt,
                        "is_company": t_rule["is_company"],
                        **pct,
                    }
                )

            for d in cfg.get("date_shapes", []):
                if int(d["slide_index"]) != s_idx:
                    continue
                sp = str(d["shape_path"]).strip()
                g = geom_map.get(sp)
                if not g or not getattr(g["shape"], "has_text_frame", False):
                    continue
                _mark_consumed(sp)
                val_txt = g["shape"].text_frame.text.strip()
                pct = _to_pct(g["left"], g["top"], g["width"], g["height"])
                elements.append(
                    {
                        "kind": "DATE_BADGE",
                        "shape_key": sp,
                        "text": val_txt,
                        **pct,
                    }
                )

            for path, g in geom_map.items():
                if path in consumed_paths or g["is_group"]:
                    continue
                sh = g["shape"]
                if not getattr(sh, "has_text_frame", False):
                    continue
                txt = sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
                if not txt:
                    continue
                if len(txt) <= 1 and not txt.isdigit():
                    continue
                pct = _to_pct(g["left"], g["top"], g["width"], g["height"])
                if pct["width_pct"] < 2.5 or pct["height_pct"] < 2.0:
                    continue

                if pct["top_pct"] < 12.0 and pct["left_pct"] < 65.0:
                    elements.append(
                        {
                            "kind": "SLIDE_TITLE",
                            "shape_key": path,
                            "text": txt,
                            **pct,
                        }
                    )
                elif pct["top_pct"] > 90.0:
                    elements.append(
                        {
                            "kind": "FOOTER_NOTE",
                            "shape_key": path,
                            "text": txt,
                            **pct,
                        }
                    )
                else:
                    area_part, mgr_part = cls.split_area_and_manager(sh)
                    elements.append(
                        {
                            "kind": "INSTITUTIONAL_BOX",
                            "shape_key": path,
                            "area_name": area_part or txt,
                            "manager": mgr_part,
                            "text": txt,
                            **pct,
                        }
                    )

            slides_data.append(
                {
                    "slide_number": s_idx + 1,
                    "elements": elements,
                    "connectors": [],
                }
            )

        return slides_data


PowerPointEngine = PptxSurgicalEngine