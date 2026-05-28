import argparse
from dataclasses import dataclass
from typing import Dict, List, Set

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ── Colores ──────────────────────────────────────────────
AZUL_OSCURO = RGBColor(0x1F, 0x4E, 0x79)
VERDE = RGBColor(0x17, 0x86, 0x4E)
AMARILLO = RGBColor(0xF1, 0xC4, 0x0F)
NARANJO = RGBColor(0xE6, 0x7E, 0x22)
GRIS_CLARO = RGBColor(0xF2, 0xF2, 0xF2)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
GRIS_TEXTO = RGBColor(0x44, 0x44, 0x44)


@dataclass
class Recomendacion:
    tecnologia: str
    categoria: str
    prioridad: float
    nivel: str
    apps: int
    relacionadas: str
    afinidad: str
    accion: str


def semaforo_afinidad(conocidas_cat: int, total_cat: int) -> str:
    ratio = conocidas_cat / max(total_cat, 1)
    if conocidas_cat >= 2 or ratio >= 0.34:
        return "🟢 Alta"
    if conocidas_cat == 1 or ratio >= 0.12:
        return "🟡 Media"
    return "🔴 Baja"


def add_rect(slide, l, t, w, h, fill_rgb):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    shape.line.fill.background()
    return shape


def add_text(
    slide,
    text,
    l,
    t,
    w,
    h,
    size=10,
    bold=False,
    color=RGBColor(0, 0, 0),
    align=PP_ALIGN.LEFT,
):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txb


def cargar_datos(archivo: str):
    raw_matriz = pd.read_excel(archivo, sheet_name="Matriz", header=None)
    categorias_row = raw_matriz.iloc[1, 2:].ffill().tolist()
    tecnologias_nombres = raw_matriz.iloc[2, 2:].tolist()
    tech_categoria = {t: c for t, c in zip(tecnologias_nombres, categorias_row)}

    matriz = pd.read_excel(archivo, sheet_name="Matriz", header=2, index_col=1)
    matriz = matriz.drop(columns=[matriz.columns[0]], errors="ignore").dropna(how="all").fillna(0)

    personas = pd.read_excel(archivo, sheet_name="MtzPerson", header=2, index_col=1)
    personas = personas.drop(columns=[personas.columns[0]], errors="ignore").dropna(how="all").fillna(0)

    tecnologias = matriz.columns.intersection(personas.columns)
    matriz = matriz[tecnologias]
    personas = personas[tecnologias]

    techs_relevantes = set(tecnologias[(matriz == 1).any(axis=0)])

    return matriz, personas, tecnologias, techs_relevantes, tech_categoria


def accion_por_categoria(categoria: str) -> str:
    c = str(categoria).lower()
    if "lenguaje" in c:
        return "Curso avanzado + mini proyecto"
    if "datos" in c or "base de datos" in c:
        return "Laboratorio de datos aplicado"
    if "seguridad" in c:
        return "Secure coding + hardening"
    if "integr" in c or "api" in c:
        return "Taller de integración end-to-end"
    if "cloud" in c or "infra" in c or "plataforma" in c:
        return "Ruta hands-on con sandbox"
    return "Curso guiado + mentoring"


def proponer_para_persona(
    persona_row: pd.Series,
    tecnologias: pd.Index,
    techs_relevantes: Set,
    tech_categoria: Dict,
    demanda_apps: Dict,
) -> List[Recomendacion]:
    conocidas = set(tecnologias[persona_row == 1]) & techs_relevantes
    gaps = list(techs_relevantes - conocidas)

    if not gaps:
        return []

    max_demanda = max(demanda_apps.values()) if demanda_apps else 1

    conocidas_por_cat = {}
    total_por_cat = {}

    for t in techs_relevantes:
        cat = str(tech_categoria.get(t, "Otros")).strip()
        total_por_cat[cat] = total_por_cat.get(cat, 0) + 1

    for t in conocidas:
        cat = str(tech_categoria.get(t, "Otros")).strip()
        conocidas_por_cat.setdefault(cat, set()).add(t)

    recs = []
    for t in gaps:
        cat = str(tech_categoria.get(t, "Otros")).strip()
        apps = int(demanda_apps.get(t, 0))
        score_demanda = apps / max_demanda
        conocidas_cat = len(conocidas_por_cat.get(cat, set()))
        total_cat = max(total_por_cat.get(cat, 1), 1)
        score_afinidad = conocidas_cat / total_cat
        prioridad = (0.6 * score_demanda) + (0.4 * score_afinidad)

        if conocidas_cat >= 2:
            nivel = "Intermedio"
        elif conocidas_cat == 1:
            nivel = "Básico-Intermedio"
        else:
            nivel = "Fundamentos"

        afinidad = semaforo_afinidad(conocidas_cat, total_cat)

        rel = sorted([str(x).strip() for x in conocidas_por_cat.get(cat, set())])
        relacionadas = ", ".join(rel[:3]) if rel else "Sin base directa"

        recs.append(
            Recomendacion(
                tecnologia=str(t).strip(),
                categoria=cat,
                prioridad=round(prioridad, 4),
                nivel=nivel,
                apps=apps,
                relacionadas=relacionadas,
                afinidad=afinidad,
                accion=accion_por_categoria(cat),
            )
        )

    recs.sort(key=lambda r: (-r.prioridad, -r.apps, r.tecnologia.lower()))
    return recs


def resumen_fase_texto(recs_fase: List[Recomendacion], limite: int = 4) -> str:
    if not recs_fase:
        return "Sin recomendaciones en esta fase"
    top = recs_fase[:limite]
    return "\n".join(
        (
            f"• {r.tecnologia} ({r.categoria}) | {r.afinidad} | Nivel: {r.nivel} | Apps: {r.apps}\n"
            f"  ↳ Intersección base: {r.relacionadas}"
        )
        for r in top
    )


def generar_presentacion(archivo_entrada: str, archivo_salida: str):
    matriz, personas, tecnologias, techs_relevantes, tech_categoria = cargar_datos(archivo_entrada)

    demanda_apps = {
        t: int((pd.to_numeric(matriz[t], errors="coerce").fillna(0) == 1).sum())
        for t in tecnologias
        if t in techs_relevantes
    }

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Portada
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, 13.33, 7.5, AZUL_OSCURO)
    add_text(s, "Propuesta de Capacitación", 1, 2.0, 11, 1.2, size=40, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
    add_text(
        s,
        "Plan personalizado por persona (Quick Wins, Consolidación y Expansión)",
        1,
        3.3,
        11,
        0.8,
        size=17,
        color=RGBColor(0xBD, 0xD7, 0xEE),
        align=PP_ALIGN.CENTER,
    )
    add_text(s, "Banco Santander Chile · 2026", 1, 4.2, 11, 0.5, size=14, color=RGBColor(0xBD, 0xD7, 0xEE), align=PP_ALIGN.CENTER)

    for persona_nombre, persona_row in personas.iterrows():
        recs = proponer_para_persona(
            persona_row=persona_row,
            tecnologias=tecnologias,
            techs_relevantes=techs_relevantes,
            tech_categoria=tech_categoria,
            demanda_apps=demanda_apps,
        )

        f1 = recs[:6]
        f2 = recs[6:12]
        f3 = recs[12:20]

        conocidas = set(tecnologias[persona_row == 1]) & techs_relevantes
        gaps = set(techs_relevantes) - conocidas
        conocidas_ordenadas = sorted([str(t).strip() for t in conocidas], key=str.lower)
        conocidas_preview = " | ".join(conocidas_ordenadas[:10]) if conocidas_ordenadas else "Sin base conocida relevante"

        conocidas_por_cat = {}
        for t in conocidas:
            cat = str(tech_categoria.get(t, "Otros")).strip()
            conocidas_por_cat[cat] = conocidas_por_cat.get(cat, 0) + 1
        top_cat = sorted(conocidas_por_cat.items(), key=lambda x: (-x[1], x[0].lower()))
        top_cat_txt = " | ".join([f"{c} ({n})" for c, n in top_cat[:4]]) if top_cat else "Sin categorías destacadas"

        s = prs.slides.add_slide(blank)
        add_rect(s, 0, 0, 13.33, 7.5, GRIS_CLARO)
        add_rect(s, 0, 0, 13.33, 1.0, AZUL_OSCURO)

        add_text(s, str(persona_nombre), 0.3, 0.08, 8.8, 0.55, size=22, bold=True, color=BLANCO)
        add_text(
            s,
            f"Conocidas: {len(conocidas)} | Brechas: {len(gaps)} | Recomendadas: {len(recs)}",
            0.3,
            0.62,
            10,
            0.3,
            size=11,
            color=RGBColor(0xBD, 0xD7, 0xEE),
        )

        # Fase headers
        add_rect(s, 0.2, 1.15, 4.25, 0.35, VERDE)
        add_text(s, "Fase 1 - Quick Wins", 0.3, 1.19, 4.05, 0.3, size=10, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

        add_rect(s, 4.55, 1.15, 4.25, 0.35, AMARILLO)
        add_text(s, "Fase 2 - Consolidación", 4.65, 1.19, 4.05, 0.3, size=10, bold=True, color=RGBColor(0x33, 0x33, 0x33), align=PP_ALIGN.CENTER)

        add_rect(s, 8.9, 1.15, 4.25, 0.35, NARANJO)
        add_text(s, "Fase 3 - Expansión", 9.0, 1.19, 4.05, 0.3, size=10, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

        # Fase bodies
        add_rect(s, 0.2, 1.52, 4.25, 4.65, BLANCO)
        add_rect(s, 4.55, 1.52, 4.25, 4.65, BLANCO)
        add_rect(s, 8.9, 1.52, 4.25, 4.65, BLANCO)

        add_text(s, resumen_fase_texto(f1, limite=5), 0.3, 1.62, 4.05, 4.45, size=9, color=GRIS_TEXTO)
        add_text(s, resumen_fase_texto(f2, limite=5), 4.65, 1.62, 4.05, 4.45, size=9, color=GRIS_TEXTO)
        add_text(s, resumen_fase_texto(f3, limite=5), 9.0, 1.62, 4.05, 4.45, size=9, color=GRIS_TEXTO)

        # Bloque inferior: base conocida + foco sugerido
        add_rect(s, 0.2, 6.15, 12.95, 1.15, RGBColor(0xEB, 0xF5, 0xFB))
        foco = recs[0] if recs else None
        if foco:
            foco_txt = (
                f"Foco inmediato sugerido: {foco.tecnologia} ({foco.categoria}) | "
                f"{foco.afinidad} | Nivel: {foco.nivel} | Apps: {foco.apps} | Intersección: {foco.relacionadas} | Acción: {foco.accion}"
            )
        else:
            foco_txt = "Sin brechas detectadas en tecnologías relevantes del banco."

        add_text(
            s,
            f"Base conocida destacada: {top_cat_txt}",
            0.35,
            6.22,
            12.6,
            0.28,
            size=8,
            bold=True,
            color=RGBColor(0x1D, 0x4E, 0x2E),
        )
        add_text(
            s,
            f"Tecnologías conocidas (top): {conocidas_preview}",
            0.35,
            6.50,
            12.6,
            0.26,
            size=8,
            bold=False,
            color=RGBColor(0x2D, 0x34, 0x36),
        )
        add_text(s, foco_txt, 0.35, 6.78, 12.6, 0.45, size=9, bold=True, color=RGBColor(0x1B, 0x4F, 0x72))

    output = archivo_salida
    try:
        prs.save(output)
    except PermissionError:
        output = "PropuestaCapacitacion_actualizada.pptx"
        prs.save(output)

    print(f"Presentación generada: {output} ({len(personas)} personas)")


def main():
    parser = argparse.ArgumentParser(description="Genera propuesta de capacitación personalizada en formato PPT")
    parser.add_argument("--input", default="AplicacionesBanco.xlsx", help="Excel de entrada")
    parser.add_argument("--output", default="PropuestaCapacitacion.pptx", help="PPT de salida")
    args = parser.parse_args()
    generar_presentacion(args.input, args.output)


if __name__ == "__main__":
    main()
