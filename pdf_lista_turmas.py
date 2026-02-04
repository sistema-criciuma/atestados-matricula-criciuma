from __future__ import annotations

import re
from datetime import datetime
from io import BytesIO
from typing import Iterable

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


def _format_date_br_any(v: str) -> str:
    s = str(v or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        y, m, d = s.split("-")
        return f"{d}/{m}/{y}"
    return s


def _pick_col(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    return None


def generate_lista_turmas_pdf(
    df: pd.DataFrame,
    escola: str,
    emitted_dt: datetime,
) -> bytes:
    styles = getSampleStyleSheet()
    s_title = styles["Heading3"]
    s_norm = styles["Normal"]

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    # colunas opcionais (tolerantes)
    c_genero = _pick_col(df, ["Gênero", "Genero", "Sexo"])
    c_nasc = _pick_col(df, ["Data de nascimento", "Data de Nascimento", "Nascimento", "Data nascimento"])

    # normaliza campos base
    df2 = df.copy()
    df2["ID_show"] = df2["ID_norm"]
    df2["INEP_show"] = df2["INEP_norm"]
    df2["Aluno_show"] = df2["Nome"].astype(str).str.strip()
    df2["Genero_show"] = df2[c_genero].astype(str).str.strip() if c_genero else ""
    df2["Nasc_show"] = df2[c_nasc].apply(_format_date_br_any) if c_nasc else ""

    # agrupamento por Curso + Série + Turma (inclui Turno no título para não misturar)
    grp_cols = ["Curso", "Série", "Turma", "Turno"]
    for c in grp_cols:
        if c not in df2.columns:
            df2[c] = ""

    groups = df2.groupby(grp_cols, dropna=False, sort=True)

    story = []
    for (curso, serie, turma, turno), g in groups:
        curso = str(curso or "").strip()
        serie = str(serie or "").strip()
        turma = str(turma or "").strip()
        turno = str(turno or "").strip()

        header_lines = [
            f"Escola: {escola}",
            f"Curso: {curso} | Série: {serie} | Turma: {turma} | Turno: {turno}",
            f"Emitido em: {emitted_dt.strftime('%d/%m/%Y %H:%M')}",
        ]
        story.append(Paragraph("<br/>".join(header_lines), s_norm))
        story.append(Spacer(1, 6 * mm))

        data = [["ID", "INEP", "Aluno", "Gênero", "Nascimento"]]
        g_sorted = g.sort_values(["Aluno_show", "ID_show"], ascending=True)

        for _, r in g_sorted.iterrows():
            data.append([
                str(r.get("ID_show", "")),
                str(r.get("INEP_show", "")),
                str(r.get("Aluno_show", "")),
                str(r.get("Genero_show", "")),
                str(r.get("Nasc_show", "")),
            ])

        # larguras ajustadas para A4
        col_widths = [18 * mm, 28 * mm, 85 * mm, 20 * mm, 28 * mm]

        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
        ]))

        story.append(tbl)
        story.append(PageBreak())

    # remove última PageBreak
    if story and isinstance(story[-1], PageBreak):
        story = story[:-1]

    doc.build(story)
    return buf.getvalue()

