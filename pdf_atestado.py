from __future__ import annotations

import os
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO
from typing import Dict, Optional, Any

from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Frame


_PT_MONTHS = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}


def date_extenso(dt: datetime) -> str:
    return f"{dt.day} de {_PT_MONTHS[dt.month]} de {dt.year}"


def normalize_intlike(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""

    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]

    if "e" in s.lower():
        try:
            d = Decimal(s)
            return str(int(d))
        except (InvalidOperation, ValueError):
            pass

    digits = re.sub(r"\D", "", s)
    return digits if digits else s


def excel_serial_to_iso(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""

    try:
        if isinstance(value, (int, float)):
            serial = float(value)
        else:
            serial = float(s) if re.fullmatch(r"\d+(\.\d+)?", s) else None
        if serial is not None and serial >= 20000:
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=int(serial))
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    return s


def format_date_br(date_value: Any) -> str:
    s = excel_serial_to_iso(date_value)
    s = (s or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        y, m, d = s.split("-")
        return f"{d}/{m}/{y}"
    return s


@dataclass
class AtestadoData:
    ano: str
    escola: str
    nome: str
    nome_mae: str
    turma: str
    serie: str
    curso: str
    turno: str
    data_matricula: str
    id_aluno: str
    inep_aluno: str


LAYOUT = {
    "margin_x_mm": 18.0,
    "logo_x_mm": 18.0,
    "logo_y_from_top_mm": 1.5,
    "logo_w_mm": 40.0,
    "logo_h_mm": 40.0,
    "right_block_x_from_right_mm": 18.0,
    "right_block_y_from_top_mm": 14.0,
    "right_block_line_gap_mm": 5.0,
    "header_center_y1_from_top_mm": 18.0,
    "header_center_y2_from_top_mm": 23.0,
    "header_center_y3_from_top_mm": 29.0,
    "header_center_y4_from_top_mm": 35.0,
    "header_center_y5_from_top_mm": 40.0,
    "emitido_y_from_top_mm": 52.0,
    "separator_y_from_top_mm": 58.0,
    "title_y_from_top_mm": 74.0,
    "title_font": 20.0,
    "body_frame_top_from_top_mm": 92.0,
    "body_frame_h_mm": 62.0,
    "body_font": 12.0,
    "body_leading": 16.0,
    "fields_y_from_top_mm": 130.0,
    "fields_line_gap_mm": 6.0,
    "valid_msg_y_from_top_mm": 200.0,
    "city_date_y_from_bottom_mm": 70.0,
    "sign_line_y_from_bottom_mm": 28.0,
    "sign_text_y_from_bottom_mm": 22.0,
}


def generate_atestado_pdf(
    data: AtestadoData,
    emitted_dt: datetime,
    city_uf: str = "Criciúma / SC",
    logo_path: Optional[str] = None,
    school_meta: Optional[Dict[str, str]] = None,
) -> bytes:
    school_meta = school_meta or {}

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4

    def y_from_top(mm_value: float) -> float:
        return H - (mm_value * mm)

    margin_x = LAYOUT["margin_x_mm"] * mm

    if logo_path and os.path.exists(logo_path):
        x = LAYOUT["logo_x_mm"] * mm
        y = H - (LAYOUT["logo_y_from_top_mm"] * mm) - (LAYOUT["logo_h_mm"] * mm)
        c.drawImage(
            logo_path,
            x, y,
            width=LAYOUT["logo_w_mm"] * mm,
            height=LAYOUT["logo_h_mm"] * mm,
            preserveAspectRatio=True,
            mask="auto",
        )

    right_x = W - (LAYOUT["right_block_x_from_right_mm"] * mm)
    right_y = H - (LAYOUT["right_block_y_from_top_mm"] * mm)
    gap = LAYOUT["right_block_line_gap_mm"] * mm

    c.setFont("Helvetica", 10)

    fone = (school_meta.get("Fone", "") or "").strip()
    inep_escola = (school_meta.get("INEP", "") or "").strip()
    email = (school_meta.get("Email", "") or "").strip()

    if fone:
        c.drawRightString(right_x, right_y, f"Fone: {fone}")
        right_y -= gap

    c.drawRightString(right_x, right_y, f"Ano Letivo: {data.ano}")
    right_y -= gap

    if inep_escola:
        c.drawRightString(right_x, right_y, f"INEP: {inep_escola}")
        right_y -= gap

    if email:
        c.drawRightString(right_x, right_y, f"E-mail: {email}")
        right_y -= gap

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(W / 2, y_from_top(LAYOUT["header_center_y1_from_top_mm"]), "PREFEITURA MUNICIPAL DE CRICIÚMA")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, y_from_top(LAYOUT["header_center_y2_from_top_mm"]), "SECRETARIA MUNICIPAL DE EDUCAÇÃO")
    c.drawCentredString(W / 2, y_from_top(LAYOUT["header_center_y3_from_top_mm"]), (data.escola or "").strip())

    c.setFont("Helvetica-Bold", 10)
    end1 = (school_meta.get("EnderecoLinha1", "") or "").strip()
    end2 = (school_meta.get("EnderecoLinha2", "") or "").strip()
    if end1:
        c.drawCentredString(W / 2, y_from_top(LAYOUT["header_center_y4_from_top_mm"]), end1)
    if end2:
        c.drawCentredString(W / 2, y_from_top(LAYOUT["header_center_y5_from_top_mm"]), end2)

    c.setFont("Helvetica", 10)
    c.drawRightString(
        right_x,
        y_from_top(LAYOUT["emitido_y_from_top_mm"]),
        f"Emitido em: {emitted_dt.strftime('%d/%m/%Y %H:%M')}"
    )

    sep_y = y_from_top(LAYOUT["separator_y_from_top_mm"])
    c.line(margin_x, sep_y, W - margin_x, sep_y)

    c.setFont("Helvetica-Bold", LAYOUT["title_font"])
    c.drawCentredString(W / 2, y_from_top(LAYOUT["title_y_from_top_mm"]), "ATESTADO DE MATRÍCULA")

    nome_up = (data.nome or "").strip().upper()
    mae_up = (data.nome_mae or "").strip().upper()

    texto = (
        f"Atestamos para os devidos fins que, o(a) aluno(a) {nome_up}, "
        f"filho(a) de {mae_up}, está regularmente matriculado no ano de {data.ano} "
        f"na(s) turma(s) {data.turma} da(s) série(s) {data.serie} do(s) curso(s) {data.curso} "
        f"no(s) período(s) {data.turno}."
    )

    style = ParagraphStyle(
        "Body",
        fontName="Helvetica",
        fontSize=LAYOUT["body_font"],
        leading=LAYOUT["body_leading"],
        alignment=TA_JUSTIFY,
    )

    p = Paragraph(texto, style)
    frame_top = y_from_top(LAYOUT["body_frame_top_from_top_mm"])
    frame_h = LAYOUT["body_frame_h_mm"] * mm
    frame = Frame(margin_x, frame_top - frame_h, W - 2 * margin_x, frame_h, showBoundary=0)
    frame.addFromList([p], c)

    y_fields = y_from_top(LAYOUT["fields_y_from_top_mm"])
    line_gap = LAYOUT["fields_line_gap_mm"] * mm

    c.setFont("Helvetica", 12)
    c.drawString(margin_x, y_fields, f"Data da matrícula: {format_date_br(data.data_matricula)}")
    y_fields -= line_gap * 1.2
    c.drawString(margin_x, y_fields, f"Código do aluno: {normalize_intlike(data.id_aluno)}")
    y_fields -= line_gap * 1.2
    c.drawString(margin_x, y_fields, f"Código nacional (INEP): {normalize_intlike(data.inep_aluno)}")

    c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, y_from_top(LAYOUT["valid_msg_y_from_top_mm"]), "Essa informação é verdadeira na data de emissão.")

    c.setFont("Helvetica", 12)
    c.drawRightString(W - margin_x, LAYOUT["city_date_y_from_bottom_mm"] * mm, f"{city_uf}, {date_extenso(emitted_dt)}.")

    sign_y = LAYOUT["sign_line_y_from_bottom_mm"] * mm
    c.line(W / 2 - 55 * mm, sign_y, W / 2 + 55 * mm, sign_y)
    c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, LAYOUT["sign_text_y_from_bottom_mm"] * mm, "ASSINATURA")

    c.showPage()
    c.save()
    return buf.getvalue()
