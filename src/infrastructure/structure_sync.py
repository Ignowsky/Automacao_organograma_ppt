"""
Sincronização dinâmica da estrutura do PowerPoint com a base (Databricks).

Em vez de depender de um slide_mapping.json "congelado" (com extintas/clonadas escritas
à mão), este módulo compara, A CADA EXECUÇÃO, o que existe no PPT anexado com o que existe
na base e devolve uma configuração no MESMO formato do JSON antigo:

  * unit_boxes        -> caixas que continuam existindo (casadas com a base)
                         + caixas sem correspondência na base marcadas com force_remove
  * cloned_new_boxes  -> unidades novas na base que ainda não têm caixa no slide
  * extinct_boxes     -> sempre vazio (a lista manual foi substituída pelo cálculo)

O JSON continua sendo usado apenas como "dica" de layout (caminhos dos selos de total,
subtotal e data), então o restante do pipeline fica intacto.
Nada é gravado em disco (Vercel é read-only): tudo acontece em memória.
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

# Só caixas "descobertas" (não presentes no JSON) com estes prefixos são removidas
# automaticamente quando não existem mais na base; as demais são mantidas e avisadas.
_UNIT_PREFIXES = ("gerencia", "superintend", "diretoria", "coordenacao", "unidade")
_ROLE_ORDER = {
    "CONSELHO_SUPERIOR": 0,
    "DIRETORIA_CABECA": 1,
    "DIRETORIA_VINCULADA": 2,
    "SUPERINTENDENCIA_TITULAR": 3,
    "GERENCIA_UNIDADE": 4,
}
_COLUMN_GAP_EMU = 120000


def _text(sh: Any) -> str:
    if sh is None or not getattr(sh, "has_text_frame", False):
        return ""
    return sh.text_frame.text.strip().replace("\x0b", " ").replace("\n", " ")


def _rec_key(r: HeadcountRecord) -> Tuple[str, str, str]:
    return (r.segmento_des, r.area_des, r.unidade_des)


class StructureSyncService:
    FUZZY_THRESHOLD = 0.97
    HINT_THRESHOLD = 0.80

    # ------------------------------------------------------------------ público
    @classmethod
    def reconcile(
        cls,
        prs: Any,
        records: List[HeadcountRecord],
        base_config: Optional[Dict[str, Any]],
    ) -> Dict[str, Any]:
        base = base_config or {}
        notes: List[str] = []
        active = [r for r in records if r.headcount > 0]
        by_tuple = {_rec_key(r): r for r in active}
        seg_names = sorted({r.segmento_des for r in active})

        seg_by_slide = cls._segment_by_slide(prs, base, seg_names)
        slide_by_seg: Dict[str, int] = {}
        for s_idx, seg in seg_by_slide.items():
            slide_by_seg.setdefault(seg, s_idx)

        candidates = cls._collect_candidates(prs, base, seg_by_slide)

        # ---------- casamento caixa <-> registro da base
        used: Set[Tuple[str, str, str]] = set()
        matched: List[Tuple[Dict[str, Any], HeadcountRecord, str]] = []
        pending: List[Dict[str, Any]] = []

        # passo 1: tupla exata (segmento, área, unidade) vinda do JSON
        for c in candidates:
            u = c["base"]
            if u is not None:
                key = (u["segmento"], u["area"], u["unidade"])
                if key in by_tuple and key not in used:
                    used.add(key)
                    matched.append((c, by_tuple[key], "exato"))
                    continue
            pending.append(c)

        # passo 2: nome normalizado igual (cobre mudança de área / caixas descobertas)
        still: List[Dict[str, Any]] = []
        for c in pending:
            slide = prs.slides[c["slide_index"]]
            name_txt = cls._box_name(slide, c)
            seg = seg_by_slide.get(c["slide_index"], "")
            pool = [r for r in active if r.segmento_des == seg and _rec_key(r) not in used]
            wanted = {normalize_text_key(name_txt)}
            if c["base"] is not None:
                wanted.add(normalize_text_key(c["base"]["unidade"]))
                if c["base"].get("rename_to"):
                    wanted.add(normalize_text_key(c["base"]["rename_to"]))
            wanted.discard("")
            hits = [r for r in pool if normalize_text_key(r.unidade_des) in wanted]
            if c["base"] is not None:
                hits.sort(key=lambda r: r.area_des != c["base"]["area"])
            if hits:
                used.add(_rec_key(hits[0]))
                matched.append((c, hits[0], "nome"))
            else:
                still.append(c)

        # passo 3: nome QUASE idêntico (acento/pontuação/abreviação mínima). Muito restrito de propósito:
        # "Gerência X2" vs "Gerência X3" são unidades diferentes e NÃO podem ser casadas.
        unmatched: List[Dict[str, Any]] = []
        for c in still:
            slide = prs.slides[c["slide_index"]]
            box_name = cls._box_name(slide, c)
            name_n = normalize_text_key(box_name)
            seg = seg_by_slide.get(c["slide_index"], "")
            pool = [r for r in active if r.segmento_des == seg and _rec_key(r) not in used]
            best, best_ratio = None, 0.0
            for r in pool:
                ratio = difflib.SequenceMatcher(None, name_n, normalize_text_key(r.unidade_des)).ratio()
                if ratio > best_ratio:
                    best, best_ratio = r, ratio
            same_digits = best is not None and (
                [ch for ch in name_n if ch.isdigit()] == [ch for ch in normalize_text_key(best.unidade_des) if ch.isdigit()]
            )
            if best is not None and best_ratio >= cls.FUZZY_THRESHOLD and same_digits:
                used.add(_rec_key(best))
                matched.append((c, best, "parecido"))
                notes.append(
                    f"Slide {c['slide_index'] + 1}: caixa \"{box_name}\" casada por similaridade "
                    f"({best_ratio:.0%}) com \"{best.unidade_des}\"."
                )
            else:
                if best is not None and best_ratio >= cls.HINT_THRESHOLD:
                    notes.append(
                        f"ℹ Slide {c['slide_index'] + 1}: \"{box_name}\" saiu da base e \"{best.unidade_des}\" é nova "
                        f"({best_ratio:.0%} de similaridade). Se for renomeação, a caixa foi recriada com "
                        f"\"(Definir Gestor)\"; confira o gestor."
                    )
                unmatched.append(c)

        # ---------- monta unit_boxes (mantidas + a remover)
        unit_boxes: List[Dict[str, Any]] = []
        for c, r, how in matched:
            entry = dict(c["base"]) if c["base"] is not None else {}
            entry.update(
                {
                    "slide_index": c["slide_index"],
                    "name_path": c["name_path"],
                    "badge_path": c["badge_path"],
                    "segmento": r.segmento_des,
                    "area": r.area_des,
                    "unidade": r.unidade_des,
                    "sync_origin": c["origin"],
                }
            )
            unit_boxes.append(entry)

        removed_count = 0
        for c in unmatched:
            slide = prs.slides[c["slide_index"]]
            name_txt = cls._box_name(slide, c)
            looks_like_unit = normalize_text_key(name_txt).startswith(_UNIT_PREFIXES)
            if c["origin"] == "descoberta" and not looks_like_unit:
                notes.append(
                    f"Slide {c['slide_index'] + 1}: caixa \"{name_txt}\" não reconhecida como unidade; mantida sem alteração."
                )
                continue
            u = c["base"] or {}
            unit_boxes.append(
                {
                    "slide_index": c["slide_index"],
                    "name_path": c["name_path"],
                    "badge_path": c["badge_path"],
                    "segmento": u.get("segmento") or seg_by_slide.get(c["slide_index"], ""),
                    "area": u.get("area") or "(sem registro na base)",
                    "unidade": u.get("unidade") or name_txt,
                    "force_remove": True,
                    "sync_origin": c["origin"],
                }
            )
            removed_count += 1

        # ---------- unidades novas na base -> clonar
        cloned = cls._plan_clones(prs, active, used, matched, seg_by_slide, slide_by_seg, notes)

        notes.insert(
            0,
            f"Estrutura sincronizada com a base: {len(matched)} caixas mantidas, "
            f"{removed_count} removidas (sem registro na base), {len(cloned)} criadas (novas na base).",
        )

        cfg = {k: v for k, v in base.items() if k not in ("unit_boxes", "cloned_new_boxes", "extinct_boxes")}
        cfg["unit_boxes"] = unit_boxes
        cfg["cloned_new_boxes"] = cloned
        cfg["extinct_boxes"] = []
        cfg["_sync_notes"] = notes
        return cfg

    # --------------------------------------------------------------- internos
    @staticmethod
    def _box_name(slide: Any, cand: Dict[str, Any]) -> str:
        sh = PptxSurgicalEngine.find_shape_by_path(slide, cand["name_path"])
        return PptxSurgicalEngine.split_area_and_manager(sh)[0]

    @classmethod
    def _segment_by_slide(cls, prs: Any, base: Dict[str, Any], seg_names: List[str]) -> Dict[int, str]:
        """slide_index -> segmento, apenas para slides de detalhe (1 segmento por slide)."""
        per_slide: Dict[int, Set[str]] = defaultdict(set)
        for s in base.get("segment_total_shapes", []):
            per_slide[int(s["slide_index"])].add(s["segmento"])
        if per_slide:
            return {i: next(iter(v)) for i, v in per_slide.items() if len(v) == 1}

        # sem JSON: tenta descobrir pelo texto do slide
        out: Dict[int, str] = {}
        for s_idx, slide in enumerate(prs.slides):
            geom = PptxSurgicalEngine._collect_slide_geometry(slide)
            blob = " ".join(normalize_text_key(_text(g["shape"])) for g in geom.values())
            hits = [s for s in seg_names if normalize_text_key(s) and normalize_text_key(s) in blob]
            if len(hits) == 1:
                out[s_idx] = hits[0]
        return out

    @classmethod
    def _collect_candidates(
        cls, prs: Any, base: Dict[str, Any], seg_by_slide: Dict[int, str]
    ) -> List[Dict[str, Any]]:
        cands: List[Dict[str, Any]] = []
        covered: Set[Tuple[int, str]] = set()

        for u in base.get("unit_boxes", []):
            s_idx = int(u["slide_index"])
            if s_idx >= len(prs.slides):
                continue
            slide = prs.slides[s_idx]
            if (
                PptxSurgicalEngine.find_shape_by_path(slide, u["name_path"]) is None
                or PptxSurgicalEngine.find_shape_by_path(slide, u["badge_path"]) is None
            ):
                continue
            cands.append(
                {"slide_index": s_idx, "name_path": str(u["name_path"]), "badge_path": str(u["badge_path"]),
                 "origin": "mapa", "base": u}
            )
            covered.add((s_idx, str(u["name_path"])))
            covered.add((s_idx, str(u["badge_path"])))

        reserved: Set[Tuple[int, str]] = set()
        for key in ("date_shapes", "company_total_shapes", "segment_total_shapes", "area_subtotal_shapes"):
            for e in base.get(key, []):
                reserved.add((int(e["slide_index"]), str(e["shape_path"]).strip()))

        # descoberta: grupo com exatamente 1 texto numérico (selo) + >=1 texto (nome)
        for s_idx in seg_by_slide:
            slide = prs.slides[s_idx]
            geom = PptxSurgicalEngine._collect_slide_geometry(slide)
            for path, g in geom.items():
                if not g["is_group"]:
                    continue
                kids = [
                    (f"{path}.{c.shape_id}", _text(c))
                    for c in g["shape"].shapes
                    if getattr(c, "has_text_frame", False) and _text(c)
                ]
                nums = [k for k in kids if k[1].replace(".", "").isdigit()]
                names = [k for k in kids if not k[1].replace(".", "").isdigit()]
                if len(nums) != 1 or not names:
                    continue
                badge_p = nums[0][0]
                name_p = max(names, key=lambda k: len(k[1]))[0]
                if (s_idx, name_p) in covered or (s_idx, badge_p) in covered:
                    continue
                if (s_idx, name_p) in reserved or (s_idx, badge_p) in reserved:
                    continue
                cands.append(
                    {"slide_index": s_idx, "name_path": name_p, "badge_path": badge_p,
                     "origin": "descoberta", "base": None}
                )
        return cands

    @staticmethod
    def _slide_unit_names(slide: Any) -> Set[str]:
        names: Set[str] = set()
        for g in PptxSurgicalEngine._collect_slide_geometry(slide).values():
            t = _text(g["shape"])
            if t:
                names.add(normalize_text_key(PptxSurgicalEngine.split_area_and_manager(g["shape"])[0]))
        names.discard("")
        return names

    @staticmethod
    def _is_simple_unit_group(slide: Any, top_id: str) -> bool:
        grp = PptxSurgicalEngine.find_shape_by_path(slide, top_id)
        if not isinstance(grp, GroupShape):
            return False
        return len([c for c in grp.shapes if getattr(c, "has_text_frame", False) and _text(c)]) == 2

    @classmethod
    def _plan_clones(
        cls,
        prs: Any,
        active: List[HeadcountRecord],
        used: Set[Tuple[str, str, str]],
        matched: List[Tuple[Dict[str, Any], HeadcountRecord, str]],
        seg_by_slide: Dict[int, str],
        slide_by_seg: Dict[str, int],
        notes: List[str],
    ) -> List[Dict[str, Any]]:
        new_recs = [r for r in active if _rec_key(r) not in used]
        new_recs.sort(
            key=lambda r: (
                r.segmento_des,
                _ROLE_ORDER.get(classify_organizational_node(*_rec_key(r))[0], 9),
                r.area_des,
                r.unidade_des,
            )
        )

        by_area: Dict[Tuple[str, str], List[Tuple[Dict[str, Any], str]]] = defaultdict(list)
        by_seg: Dict[str, List[Tuple[Dict[str, Any], str]]] = defaultdict(list)
        for c, r, _ in matched:
            role = classify_organizational_node(*_rec_key(r))[0]
            by_area[(r.segmento_des, r.area_des)].append((c, role))
            by_seg[r.segmento_des].append((c, role))

        out: List[Dict[str, Any]] = []
        new_area_cursor: Dict[Tuple[str, str], Dict[str, int]] = {}

        for r in new_recs:
            s_idx = slide_by_seg.get(r.segmento_des)
            if s_idx is None:
                notes.append(
                    f"⚠ \"{r.unidade_des}\" ({r.segmento_des}) é novo na base, mas não há slide desse segmento no PPT. "
                    f"Crie o slide modelo e reenvie."
                )
                continue
            slide = prs.slides[s_idx]
            if normalize_text_key(r.unidade_des) in cls._slide_unit_names(slide):
                notes.append(
                    f"ℹ \"{r.unidade_des}\" já aparece no slide {s_idx + 1}, mas não foi mapeada (caixa fora do padrão "
                    f"nome+selo). Não foi clonada para evitar duplicidade; confira o número manualmente."
                )
                continue
            role_code = classify_organizational_node(*_rec_key(r))[0]
            area_sibs = [(c, rl) for c, rl in by_area.get((r.segmento_des, r.area_des), []) if c["slide_index"] == s_idx]
            seg_sibs = [(c, rl) for c, rl in by_seg.get(r.segmento_des, []) if c["slide_index"] == s_idx]
            pool = area_sibs or seg_sibs
            if not pool:
                notes.append(
                    f"⚠ \"{r.unidade_des}\" ({r.segmento_des}) é novo na base, mas o slide {s_idx + 1} não tem nenhuma "
                    f"caixa-modelo para clonar."
                )
                continue

            def _pick(sibs: List[Tuple[Dict[str, Any], str]]) -> Optional[Dict[str, Any]]:
                simple = [(c, rl) for c, rl in sibs if cls._is_simple_unit_group(slide, c["name_path"].split(".")[0])]
                for c, rl in simple:
                    if rl == role_code:
                        return c
                if simple:
                    return simple[-1][0]
                return None

            tpl = _pick(area_sibs) or _pick(seg_sibs) or pool[-1][0]
            entry: Dict[str, Any] = {
                "slide_index": s_idx,
                "template_group_id": int(tpl["name_path"].split(".")[0]),
                "segmento": r.segmento_des,
                "area": r.area_des,
                "unidade": r.unidade_des,
                "display_name": r.unidade_des,
                "manager_placeholder": "(Definir Gestor)",
            }

            if not area_sibs:
                # área inexistente no slide: abre uma coluna nova à direita da última coluna
                cursor = new_area_cursor.get((r.segmento_des, r.area_des))
                if cursor is None:
                    geom = PptxSurgicalEngine._collect_slide_geometry(slide)
                    tops = [geom[c["name_path"].split(".")[0]] for c, _ in seg_sibs if c["name_path"].split(".")[0] in geom]
                    tpl_g = geom.get(str(entry["template_group_id"]))
                    right = max((g["left"] + g["width"] for g in tops), default=0)
                    head_tops = [
                        geom[c["name_path"].split(".")[0]]["top"]
                        for c, rl in seg_sibs
                        if rl == "SUPERINTENDENCIA_TITULAR" and c["name_path"].split(".")[0] in geom
                    ]
                    cursor = {
                        "left": int(right + _COLUMN_GAP_EMU),
                        "top": int(min(head_tops) if head_tops else (tpl_g["top"] if tpl_g else 1850000)),
                    }
                    new_area_cursor[(r.segmento_des, r.area_des)] = cursor
                    slide_w = int(prs.slide_width or 12192000)
                    tpl_w = int(tpl_g["width"]) if tpl_g else 0
                    notes.append(
                        f"Nova área \"{r.area_des}\" ({r.segmento_des}): coluna criada no slide {s_idx + 1}. "
                        f"O selo de subtotal da nova área NÃO é criado automaticamente; confira o posicionamento."
                    )
                    if cursor["left"] + tpl_w > slide_w:
                        notes.append(
                            f"⚠ Slide {s_idx + 1}: a nova coluna \"{r.area_des}\" ultrapassa a largura do slide; reorganize manualmente."
                        )
                entry["left_override"] = cursor["left"]
                entry["top_override"] = cursor["top"]
                # próximas caixas da mesma área nova ficam empilhadas abaixo
                cursor["top"] += PptxSurgicalEngine.STANDARD_UNIT_HEIGHT_EMU + PptxSurgicalEngine.DEFAULT_VERTICAL_GAP_EMU

            out.append(entry)
        return out