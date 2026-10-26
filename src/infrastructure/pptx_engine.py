import copy
from typing import Optional, Dict, Any, List
from pptx.presentation import Presentation
from pptx.slide import Slide
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.oxml.ns import qn


class PptxSurgicalEngine:
    """
    Motor cirúrgico de baixo nível para leitura, atualização de Runs, edição manual de gestores,
    clonagem/remoção de grupos com compactação de coluna e extração de Preview 16:9 no python-pptx.
    """

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
        """Retorna separadamente (nome_da_area, nome_do_gestor_com_parenteses)."""
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
        """
        Substitui o nome da área na primeira Run (<a:r>), preserva a quebra de linha (<a:br>)
        e mantém o nome do gestor intacto — a menos que 'new_manager_override' seja enviado pela Web.
        """
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

    @classmethod
    def clone_unit_group_in_column(
        cls,
        slide: Slide,
        template_group_id: int,
        new_area_name: str,
        new_headcount: int,
        manager_placeholder: str = "(Definir Gestor)",
        left_override: Optional[int] = None,
        top_override: Optional[int] = None,
    ) -> Dict[str, Any]:
        template_group = cls.find_shape_by_path(slide, str(template_group_id))
        if template_group is None or not isinstance(template_group, GroupShape):
            raise ValueError(f"Grupo base {template_group_id} não encontrado para clonagem.")

        new_el = copy.deepcopy(template_group.element)
        max_id = max(s.shape_id for s in slide.shapes)
        new_group_id = max_id + 100
        new_el.nvGrpSpPr.cNvPr.id = new_group_id
        new_el.nvGrpSpPr.cNvPr.name = f"Agrupar_Auto_{new_group_id}"

        slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
        cloned_group: GroupShape = slide.shapes[-1]

        target_left = left_override if left_override is not None else template_group.left

        # Se top_override não foi fixado, posiciona logo abaixo da última caixa de unidade daquela coluna
        if top_override is not None:
            target_top = top_override
        else:
            lowest_bottom = template_group.top + template_group.height
            for sh in slide.shapes:
                if sh.shape_id == new_group_id:
                    continue
                if isinstance(sh, GroupShape) and abs(sh.left - target_left) <= cls.COLUMN_ALIGNMENT_TOLERANCE_EMU:
                    sh_bottom = sh.top + sh.height
                    if sh_bottom > lowest_bottom and sh.top < 5600000:
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

        return {
            "new_group_id": new_group_id,
            "name_path": f"{new_group_id}.{text_shapes[0].shape_id}" if len(text_shapes) >= 2 else str(new_group_id),
            "badge_path": f"{new_group_id}.{text_shapes[1].shape_id}" if len(text_shapes) >= 2 else str(new_group_id),
            "left": cloned_group.left,
            "top": cloned_group.top,
            "area_name": new_area_name,
            "headcount": new_headcount,
        }

    @classmethod
    def remove_unit_group_in_column(
        cls,
        slide: Slide,
        shape_key: str,
        compact_column: bool = True,
    ) -> Optional[Dict[str, Any]]:
        """
        Remove fisicamente uma caixa (GroupShape ou Shape) do slide e, se compact_column=True,
        desloca as caixas inferiores da mesma coluna para cima para não deixar buraco no layout.
        """
        first_part = shape_key.split("|")[0].strip()
        root_id_str = first_part.split(".")[0].strip()
        target_root = cls.find_shape_by_path(slide, root_id_str)
        if target_root is None:
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
        elif getattr(target_root, "has_text_frame", False):
            old_text = target_root.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")

        removed_left = int(getattr(target_root, "left", 0))
        removed_top = int(getattr(target_root, "top", 0))
        removed_height = int(getattr(target_root, "height", cls.STANDARD_UNIT_HEIGHT_EMU))

        # Remove o elemento do XML do slide
        slide.shapes._spTree.remove(target_root.element)

        # Compacta a coluna subindo as caixas de gerência que estavam abaixo da caixa removida
        shifted_count = 0
        if compact_column and removed_height > 0:
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
            "removed_root_id": root_id_str,
            "old_text": old_text,
            "old_hc": old_hc,
            "left": removed_left,
            "top": removed_top,
            "shifted_count": shifted_count,
        }

    @classmethod
    def extract_slides_preview(cls, prs: Presentation) -> List[Dict[str, Any]]:
        """
        Extrai a geometria normalizada (0 a 100%) e os textos de todas as caixas dos slides
        para renderizar uma Pré-Visualização Interativa 16:9 fiel diretamente no navegador.
        """
        slide_w = float(prs.slide_width or 12192000)
        slide_h = float(prs.slide_height or 6858000)
        slides_data: List[Dict[str, Any]] = []

        for s_idx, slide in enumerate(prs.slides):
            elements: List[Dict[str, Any]] = []

            for sh in slide.shapes:
                left_pct = round((float(getattr(sh, "left", 0)) / slide_w) * 100, 2)
                top_pct = round((float(getattr(sh, "top", 0)) / slide_h) * 100, 2)
                width_pct = round((float(getattr(sh, "width", 0)) / slide_w) * 100, 2)
                height_pct = round((float(getattr(sh, "height", 0)) / slide_h) * 100, 2)

                if isinstance(sh, GroupShape):
                    text_children = [
                        c for c in sh.shapes if getattr(c, "has_text_frame", False) and c.text_frame.text.strip()
                    ]
                    if len(text_children) >= 2:
                        name_sh, badge_sh = text_children[0], text_children[1]
                        area_name, mgr_name = cls.split_area_and_manager(name_sh)
                        badge_txt = badge_sh.text_frame.text.strip()
                        shape_key = f"{sh.shape_id}.{name_sh.shape_id} | {sh.shape_id}.{badge_sh.shape_id}"
                        elements.append(
                            {
                                "kind": "UNIT_GROUP",
                                "shape_key": shape_key,
                                "group_id": sh.shape_id,
                                "left_pct": left_pct,
                                "top_pct": top_pct,
                                "width_pct": max(width_pct, 8.0),
                                "height_pct": max(height_pct, 4.5),
                                "area_name": area_name,
                                "manager": mgr_name,
                                "badge": badge_txt,
                            }
                        )
                    elif len(text_children) == 1:
                        only_sh = text_children[0]
                        txt = only_sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
                        elements.append(
                            {
                                "kind": "SINGLE_BOX",
                                "shape_key": f"{sh.shape_id}.{only_sh.shape_id}",
                                "group_id": sh.shape_id,
                                "left_pct": left_pct,
                                "top_pct": top_pct,
                                "width_pct": max(width_pct, 4.0),
                                "height_pct": max(height_pct, 3.5),
                                "text": txt,
                            }
                        )
                elif getattr(sh, "has_text_frame", False):
                    txt = sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")
                    if not txt:
                        continue
                    is_numeric = txt.isdigit()
                    elements.append(
                        {
                            "kind": "BADGE_OR_TOTAL" if is_numeric else "TEXT_SHAPE",
                            "shape_key": str(sh.shape_id),
                            "group_id": sh.shape_id,
                            "left_pct": left_pct,
                            "top_pct": top_pct,
                            "width_pct": max(width_pct, 3.2 if is_numeric else 6.0),
                            "height_pct": max(height_pct, 3.2),
                            "text": txt,
                        }
                    )

            slides_data.append(
                {
                    "slide_number": s_idx + 1,
                    "elements": elements,
                }
            )

        return slides_data