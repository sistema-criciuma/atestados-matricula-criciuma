from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Optional

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from pdf_atestado import normalize_intlike


@dataclass
class HeaderInfo:
    escola: str
    curso: str
    serie: str
    turma: str
    turno: str


def _safe_str(v) -> str:
    s = "" if v is None else str(v)
    return s.strip()


def _first_non_empty(series: pd.Series) -> str:
    for v in series.dropna().tolist():
        s = _safe_str(v)
        if s:
            return s
    return ""


def _emitido_str(emitted_dt: datetime) -> str:
    # emitted_dt já deve vir com timezone correto do app (datetime.now(TZ))
    return emitted_dt.strftime("%d/%m/%Y %H:%M")


def generate_lista_turmas_pdf(
    df: pd.DataFrame,
    *,
    escola: str,
    emitted_dt: datetime,
    logo_path: Optional[str] = None,
    city_uf: str = "Criciúma / SC",
) -> bytes:
    """
    Gera PDF com listas por turma (uma seção por turma, com quebra de página).
    Tabela: ID, INEP, Aluno, Livre
    Inclui totalizador por turma.
    Inclui logo em todas as páginas (se existir).
    """

    if df is None or df.empty:
        return b""

    # Normalizações mínimas esperadas pelo app, mas garantimos aqui também
    if "ID_norm" not in df.columns and "ID Aluno" in df.columns:
        df["ID_norm"] = df["ID Aluno"].apply(normalize_intlike)
    if "INEP_norm" not in df.columns and "Código INEP (Aluno)" in df.columns:
        df["INEP_norm"] = df["Código INEP (Aluno)"].apply(normalize_intlike)

    # Filtro de escola (defensivo)
    escola_up = _safe_str(escola).upper()
    if "Escola" in df.columns:
        df = df.loc[df["Escola"].astype(str).str.strip().str.upper() == escola_up].copy()

    if df.empty:
        return b""

    # Ordenação para consistência
    for col in ("Curso", "Série", "Turma", "Nome"):
        if col in df.columns:
            df[col] = df[col].astype(str)

    sort_cols = [c for c in ["Curso", "Série", "Turma", "Turno", "Nome"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols)

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        title="Lista de alunos por turma",
    )

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(
        "title",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        spaceAfter=6,
    )
    style_header = ParagraphStyle(
        "header",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=13,
        spaceAfter=2,
    )
    style_small = ParagraphStyle(
        "small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=12,
    )
    style_bold_line = ParagraphStyle(
        "boldline",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=13,
        spaceBefore=6,
        spaceAfter=0,
    )

    story = []

    # Grupos (uma página por turma)
    group_cols = [c for c in ["Curso", "Série", "Turma", "Turno"] if c in df.columns]
    if not group_cols:
        # Se por algum motivo não existir, cria um grupo único
        df["_grp"] = "TURMA"
        group_cols = ["_grp"]

    grouped = df.groupby(group_cols, dropna=False)

    for gi, (gkey, gdf) in enumerate(grouped):
        if gi > 0:
            story.append(Spacer(1, 2 * mm))
            story.append(Paragraph("<br/>", style_small))
            story.append(Spacer(1, 2 * mm))
            # quebra de página
            story.append(PageBreak())

        # Header info
        if not isinstance(gkey, tuple):
            gkey = (gkey,)

        # Monta campos do cabeçalho conforme existirem
        curso = _first_non_empty(gdf["Curso"]) if "Curso" in gdf.columns else ""
        serie = _first_non_empty(gdf["Série"]) if "Série" in gdf.columns else ""
        turma = _first_non_empty(gdf["Turma"]) if "Turma" in gdf.columns else ""
        turno = _first_non_empty(gdf["Turno"]) if "Turno" in gdf.columns else ""

        # Logo em todas as páginas (se existir)
        if logo_path and os.path.exists(logo_path):
            try:
                img = Image(logo_path)
                img.drawHeight = 18 * mm
                img.drawWidth = 18 * mm
                story.append(img)
                story.append(Spacer(1, 2 * mm))
            except Exception:
                pass

        story.append(Paragraph("LISTA DE ALUNOS POR TURMA", style_title))
        story.append(Paragraph(f"Escola: {escola_up}", style_header))
        if curso:
            story.append(Paragraph(f"Curso: {curso}", style_header))
        if serie:
            story.append(Paragraph(f"Série: {serie}", style_header))
        if turma:
            story.append(Paragraph(f"Turma: {turma}", style_header))
        if turno:
            story.append(Paragraph(f"Turno: {turno}", style_header))
        story.append(Paragraph(f"Emitido em: {_emitido_str(emitted_dt)}", style_header))
        story.append(Spacer(1, 4 * mm))

        # Tabela: ID, INEP, Nome, Livre
        header_row = ["ID", "INEP", "Aluno", "Livre"]
        rows = [header_row]

        # Linhas
        for _, r in gdf.iterrows():
            id_aluno = _safe_str(r.get("ID_norm", r.get("ID Aluno", "")))
            inep = _safe_str(r.get("INEP_norm", r.get("Código INEP (Aluno)", "")))
            nome = _safe_str(r.get("Nome", ""))

            # Usa Paragraph para permitir quebra de linha do nome
            nome_cell = Paragraph(nome, style_small)
            rows.append([id_aluno, inep, nome_cell, ""])

        # Larguras (A4 útil ~170mm com margens 20mm/20mm)
        col_widths = [15 * mm, 28 * mm, 50 * mm, 49 * mm]

        table = Table(rows, colWidths=col_widths, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 10),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )

        story.append(table)

        # Totalizador por turma (linha em negrito abaixo da tabela)
        total = int(len(gdf))
        story.append(Paragraph(f"Total de alunos na turma: {total}", style_bold_line))

        # Espaço final
        story.append(Spacer(1, 6 * mm))

    doc.build(story)
    return buf.getvalue()


# ReportLab precisa disso se você usar PageBreak como flowable
from reportlab.platypus import PageBreak  # noqa: E402

