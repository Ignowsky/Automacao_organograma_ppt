import copy
from typing import Optional, Dict, Any
from pptx.slide import Slide
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.oxml.ns import qn


class PptxSurgicalEngine:
    """
    Motor cirúrgico de baixo nível para leitura, atualização de Runs, edição manual de gestores
    e clonagem de grupos no python-pptx preservando rigorosamente o layout.
    """

    DEFAULT_VERTICAL_GAP_EMU = 95000
    STANDARD_UNIT_HEIGHT_EMU = 629312

    @staticmethod
    def find_shape_by_path(slide: Slide, shape_path: str | int) -> Optional[BaseShape]:
        parts = [int(p) for p in str(shape_path).split(".")]
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

            # Se o utilizador editou o nome do gestor na interface Web:
            if new_manager_override is not None:
                mgr_formatted = new_manager_override.strip()
                if mgr_formatted and not mgr_formatted.startswith("("):
                    mgr_formatted = f"({mgr_formatted})"
                t_mgr = mgr_r_el.find(qn("a:t"))
                if t_mgr is not None:
                    t_mgr.text = mgr_formatted
                # Limpa runs extras de sobrenome partido (ex: Biasutti)
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

        cloned_group.left = left_override if left_override is not None else template_group.left
        cloned_group.top = (
            top_override
            if top_override is not None
            else template_group.top + template_group.height + cls.DEFAULT_VERTICAL_GAP_EMU
        )
        cloned_group.width = template_group.width
        # Usa altura padrão compacta se o texto for curto (como 'Logística CNI')
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