from __future__ import annotations

import base64
import io
import os
import re
from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.parse import quote

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image as RLImage,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from streamlit_folium import st_folium
import folium
from streamlit_js_eval import get_geolocation


APP_TITLE = "Generador de Informes Institucionales"
DATA_FILE = "Datos Importantes.xlsx"
PROGRAMAS = ["GREAT", "DARE", "VIFA", "MPAS", "PSCC", "Ligas Atléticas", "TOP 20", "ESS", "LAP"]

# Actividades oficiales mostradas para el Programa de Seguridad Comunitaria y Comercial.
# Para los demás programas, la aplicación continúa tomando las actividades desde el Excel.
ACTIVIDADES_POR_PROGRAMA = {
    "PSCC": [
        "Curso básico",
        "Creación del Comité de Seguridad Comunitaria",
        "Seguimiento a Comités de Seguridad Comunitaria",
        "Capacitación en Gestión",
        "Capacitación en Participación Ciudadana",
        "Capacitación en Denuncia Comunitaria",
        "Capacitación en Cultura Preventiva",
        "Actividad de Cohesión Social",
        "Actividad de Recuperación de Espacios Públicos",
        "Otras Actividades de Seguridad Comunitaria",
        "Otros",
    ],
    "TOP 20": ["RPD", "DP", "CIR Social", "MSC", "CCCI", "MAL", "Otros"],
    "ESS": ["Reuniones", "Seguimiento", "Otros"],
    "LAP": ["Otros"],
    "Ligas Atléticas": ["Otros"],
}
PRIMARY_BLUE = "#173B67"
SECONDARY_BLUE = "#267FB8"
LIGHT_BLUE = "#EAF3F8"
DARK_TEXT = "#1D2939"


@dataclass
class Catalogos:
    territorios: pd.DataFrame
    regiones_delegaciones: pd.DataFrame
    actividades: pd.DataFrame
    top20_col: Optional[str]


def normalizar_texto(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def primer_archivo_existente(candidatos: Sequence[str]) -> Optional[Path]:
    for nombre in candidatos:
        ruta = Path(nombre)
        if ruta.exists() and ruta.is_file():
            return ruta
    return None


def encontrar_logos() -> List[Path]:
    candidatos = [
        "Logo1.jpeg", "Logo1.jpg", "Logo1.png",
        "logo1.jpeg", "logo1.jpg", "logo1.png",
        "Logo2.jpeg", "Logo2.jpg", "Logo2.png",
        "logo2.jpeg", "logo2.jpg", "logo2.png",
        "logo_msp.jpeg", "logo_msp.jpg", "logo_msp.png",
        "logo_gobierno.jpeg", "logo_gobierno.jpg", "logo_gobierno.png",
    ]
    vistos: set[str] = set()
    logos: List[Path] = []
    for nombre in candidatos:
        ruta = Path(nombre)
        clave = str(ruta.resolve()) if ruta.exists() else nombre
        if ruta.exists() and ruta.is_file() and clave not in vistos:
            logos.append(ruta)
            vistos.add(clave)
    return logos


@st.cache_data(show_spinner=False)
def cargar_catalogos(ruta_excel: str) -> Catalogos:
    df = pd.read_excel(ruta_excel, sheet_name=0, dtype=str)
    df.columns = [normalizar_texto(c) for c in df.columns]

    requeridas = {"Provincia", "Cantón", "Distritos", "Dirección Regional", "Delegación", "Actividad Realizada", "Programa"}
    faltantes = sorted(requeridas - set(df.columns))
    if faltantes:
        raise ValueError("Faltan columnas obligatorias en el Excel: " + ", ".join(faltantes))

    for col in df.columns:
        df[col] = df[col].map(normalizar_texto)

    territorios = (
        df[["Provincia", "Cantón", "Distritos"]]
        .replace("", pd.NA)
        .dropna(subset=["Provincia", "Cantón", "Distritos"])
        .drop_duplicates()
        .sort_values(["Provincia", "Cantón", "Distritos"], kind="stable")
        .reset_index(drop=True)
    )

    regiones_delegaciones = (
        df[["Dirección Regional", "Delegación"]]
        .replace("", pd.NA)
        .dropna(subset=["Dirección Regional", "Delegación"])
        .drop_duplicates()
        .reset_index(drop=True)
    )

    columnas_actividades = [c for c in ["Responde a:", "Actividad Realizada", "Programa"] if c in df.columns]
    actividades = (
        df[columnas_actividades]
        .replace("", pd.NA)
        .dropna(subset=["Actividad Realizada", "Programa"])
        .drop_duplicates()
        .reset_index(drop=True)
    )

    top20_col = next((c for c in df.columns if "top" in c.lower() and "20" in c.lower()), None)
    if top20_col:
        territorios = territorios.merge(
            df[["Provincia", "Cantón", "Distritos", top20_col]].drop_duplicates(),
            on=["Provincia", "Cantón", "Distritos"],
            how="left",
        )

    return Catalogos(
        territorios=territorios,
        regiones_delegaciones=regiones_delegaciones,
        actividades=actividades,
        top20_col=top20_col,
    )


def es_top20_automatico(catalogos: Catalogos, provincia: str, canton: str, distrito: str) -> Optional[bool]:
    if not catalogos.top20_col:
        return None
    filtro = catalogos.territorios[
        (catalogos.territorios["Provincia"] == provincia)
        & (catalogos.territorios["Cantón"] == canton)
        & (catalogos.territorios["Distritos"] == distrito)
    ]
    if filtro.empty:
        return None
    valor = normalizar_texto(filtro.iloc[0][catalogos.top20_col]).lower()
    if valor in {"sí", "si", "s", "1", "x", "true", "verdadero", "top 20", "top20"}:
        return True
    if valor in {"no", "n", "0", "false", "falso"}:
        return False
    return None


def numero_seguro(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def imagen_a_jpeg_bytes(uploaded_file: Any, max_px: int = 1800) -> bytes:
    uploaded_file.seek(0)
    img = Image.open(uploaded_file)
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((max_px, max_px))
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=86, optimize=True)
    return buffer.getvalue()


def logo_para_reportlab(path: Path, max_w: float, max_h: float) -> RLImage:
    img = Image.open(path)
    w, h = img.size
    escala = min(max_w / w, max_h / h)
    return RLImage(str(path), width=w * escala, height=h * escala)


def photo_flowable(photo_bytes: bytes, max_w: float = 16.0 * cm, max_h: float = 10.5 * cm) -> RLImage:
    img = Image.open(io.BytesIO(photo_bytes))
    w, h = img.size
    escala = min(max_w / w, max_h / h)
    bio = io.BytesIO(photo_bytes)
    return RLImage(bio, width=w * escala, height=h * escala)


def registrar_fuente() -> Tuple[str, str]:
    regular = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    try:
        if os.path.exists(regular) and os.path.exists(bold):
            pdfmetrics.registerFont(TTFont("DejaVuSans", regular))
            pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", bold))
            return "DejaVuSans", "DejaVuSans-Bold"
    except Exception:
        pass
    return "Helvetica", "Helvetica-Bold"


def encabezado_pie(canvas, doc, logos: List[Path], font_regular: str, font_bold: str):
    canvas.saveState()
    page_w, page_h = letter
    margen = 1.6 * cm

    if logos:
        x = margen
        for logo_path in logos[:2]:
            try:
                logo = logo_para_reportlab(logo_path, max_w=8.5 * cm, max_h=1.35 * cm)
                logo.drawOn(canvas, x, page_h - 1.65 * cm)
                x += logo.drawWidth + 0.35 * cm
            except Exception:
                pass

    canvas.setStrokeColor(colors.HexColor(PRIMARY_BLUE))
    canvas.setLineWidth(1.2)
    canvas.line(margen, page_h - 1.88 * cm, page_w - margen, page_h - 1.88 * cm)

    canvas.setFont(font_regular, 7.6)
    canvas.setFillColor(colors.HexColor("#667085"))
    canvas.drawString(margen, 0.78 * cm, "Ministerio de Seguridad Pública - Informe institucional de visita")
    canvas.drawRightString(page_w - margen, 0.78 * cm, f"Página {doc.page}")
    canvas.restoreState()


def construir_pdf(datos: Dict[str, Any], fotos: List[bytes], logos: List[Path]) -> bytes:
    from xml.sax.saxutils import escape

    font_regular, font_bold = registrar_fuente()
    buffer = io.BytesIO()

    dependencia = normalizar_texto(datos.get("delegacion_visitada")) or normalizar_texto(datos.get("direccion_regional")) or "Visita institucional"
    doc = BaseDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=1.65 * cm,
        leftMargin=1.65 * cm,
        topMargin=2.35 * cm,
        bottomMargin=1.45 * cm,
        title=f"Informe institucional - {dependencia}",
        author="Ministerio de Seguridad Pública",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="contenido")
    doc.addPageTemplates([PageTemplate(
        id="institucional",
        frames=[frame],
        onPage=lambda canvas, d: encabezado_pie(canvas, d, logos, font_regular, font_bold),
    )])

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloInst", parent=styles["Title"], fontName=font_bold,
        fontSize=16, leading=20, textColor=colors.HexColor(PRIMARY_BLUE), alignment=TA_CENTER, spaceAfter=8))
    styles.add(ParagraphStyle(name="SubtituloInst", parent=styles["Normal"], fontName=font_regular,
        fontSize=9.5, leading=13, textColor=colors.HexColor("#475467"), alignment=TA_CENTER, spaceAfter=14))
    styles.add(ParagraphStyle(name="Seccion", parent=styles["Heading2"], fontName=font_bold,
        fontSize=11.5, leading=14, textColor=colors.white, backColor=colors.HexColor(PRIMARY_BLUE),
        borderPadding=(5, 7, 5, 7), spaceBefore=9, spaceAfter=7))
    styles.add(ParagraphStyle(name="Texto", parent=styles["BodyText"], fontName=font_regular,
        fontSize=9.1, leading=13, textColor=colors.HexColor(DARK_TEXT), alignment=TA_JUSTIFY, spaceAfter=6))
    styles.add(ParagraphStyle(name="CeldaEtiqueta", parent=styles["BodyText"], fontName=font_bold,
        fontSize=8.2, leading=10.5, textColor=colors.HexColor(PRIMARY_BLUE)))
    styles.add(ParagraphStyle(name="CeldaValor", parent=styles["BodyText"], fontName=font_regular,
        fontSize=8.2, leading=10.5, textColor=colors.HexColor(DARK_TEXT)))
    styles.add(ParagraphStyle(name="FotoCaption", parent=styles["BodyText"], fontName=font_regular,
        fontSize=8, leading=10, textColor=colors.HexColor("#667085"), alignment=TA_CENTER, spaceAfter=8))

    def tiene_valor(valor: Any) -> bool:
        if valor is None:
            return False
        if isinstance(valor, str):
            return bool(normalizar_texto(valor))
        return True

    def texto_fecha(valor: Any, con_hora: bool = False) -> str:
        if not valor:
            return ""
        if con_hora:
            return valor.strftime("%H:%M")
        return valor.strftime("%d/%m/%Y")

    def pcelda(text: Any, style: str = "CeldaValor", raw: bool = False) -> Paragraph:
        contenido = normalizar_texto(text)
        if not raw:
            contenido = escape(contenido).replace("\n", "<br/>")
        return Paragraph(contenido, styles[style])

    def filas_validas(filas: List[Tuple[str, Any]]) -> List[Tuple[str, Any]]:
        return [(etq, val) for etq, val in filas if tiene_valor(val)]

    def tabla_datos(filas: List[Tuple[str, Any]], widths=(5.2 * cm, 11.0 * cm)) -> Optional[Table]:
        filas = filas_validas(filas)
        if not filas:
            return None
        data = [[pcelda(etq, "CeldaEtiqueta"), pcelda(val)] for etq, val in filas]
        tabla = Table(data, colWidths=list(widths), hAlign="LEFT", splitByRow=True)
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(LIGHT_BLUE)),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return tabla

    def agregar_seccion(story: List[Any], titulo: str, filas: List[Tuple[str, Any]]) -> None:
        tabla = tabla_datos(filas)
        if tabla is not None:
            story.append(Paragraph(titulo, styles["Seccion"]))
            story.append(tabla)

    story: List[Any] = [Spacer(1, 0.2 * cm), Paragraph("INFORME INSTITUCIONAL DE VISITA", styles["TituloInst"])]
    subtitulo_partes = [x for x in [normalizar_texto(datos.get("direccion_regional")), normalizar_texto(datos.get("delegacion_visitada"))] if x]
    fecha_hora = ""
    if datos.get("fecha_visita"):
        fecha_hora = f"Fecha de la visita: {texto_fecha(datos['fecha_visita'])}"
        if datos.get("hora_visita"):
            fecha_hora += f" a las {texto_fecha(datos['hora_visita'], con_hora=True)}"
    if fecha_hora:
        subtitulo_partes.append(fecha_hora)
    if subtitulo_partes:
        story.append(Paragraph("<br/>".join(escape(x) for x in subtitulo_partes), styles["SubtituloInst"]))

    intro_partes = [
        "El presente informe consolida la información registrada durante una visita institucional y constituye un respaldo para la trazabilidad de las actuaciones, la valoración técnica y el seguimiento de los acuerdos adoptados."
    ]
    if datos.get("proposito"):
        intro_partes.append(f"La visita tuvo como propósito {escape(normalizar_texto(datos['proposito']).lower())}.")
    if datos.get("programa") or datos.get("actividad"):
        intro_partes.append("La información permite documentar las acciones, líneas o actividades desarrolladas dentro del ámbito preventivo e institucional correspondiente.")
    story.append(Paragraph("1. Introducción", styles["Seccion"]))
    story.append(Paragraph(" ".join(intro_partes), styles["Texto"]))

    fecha_hora_valor = ""
    if datos.get("fecha_visita"):
        fecha_hora_valor = texto_fecha(datos["fecha_visita"])
        if datos.get("hora_visita"):
            fecha_hora_valor += f" - {texto_fecha(datos['hora_visita'], con_hora=True)}"
    agregar_seccion(story, "2. Información general de la visita", [
        ("Dirección Regional o dependencia que realiza la visita", datos.get("direccion_regional")),
        ("Delegación Policial visitada", datos.get("delegacion_visitada")),
        ("Modalidad", datos.get("modalidad")),
        ("Propósito", datos.get("proposito")),
        ("Fecha y hora", fecha_hora_valor),
        ("Persona(s) funcionaria(s) que realiza(n) la visita", datos.get("funcionarios_realizan")),
        ("Persona(s) funcionaria(s) que atiende(n) la visita", datos.get("funcionarios_atienden")),
    ])

    top20 = datos.get("es_top20")
    top20_text = "Sí" if top20 is True else "No" if top20 is False else ""
    ubicacion_filas = [
        ("Provincia", datos.get("provincia")), ("Cantón", datos.get("canton")),
        ("Distrito", datos.get("distrito")), ("Distrito perteneciente al Top 20", top20_text),
        ("Referencia del lugar", datos.get("referencia_lugar")),
    ]
    if datos.get("latitud") is not None and datos.get("longitud") is not None:
        ubicacion_filas.append(("Coordenadas", f"Latitud: {datos['latitud']:.6f} | Longitud: {datos['longitud']:.6f}"))
        ubicacion_filas.append(("Ubicación en mapa", f"https://www.openstreetmap.org/?mlat={datos['latitud']}&mlon={datos['longitud']}"))
    agregar_seccion(story, "3. Ubicación y referencia territorial", ubicacion_filas)

    agregar_seccion(story, "4. Programa y actividad valorada", [
        ("Programa Policial Preventivo y/o actividad", datos.get("programa")),
        ("Actividad evaluada o valorada", datos.get("actividad")),
        ("Marco de planificación", datos.get("responde_a")),
        ("Nombres de las acciones, líneas o actividades", datos.get("lineas_accion")),
    ])

    evidencia = datos.get("tiene_evidencia")
    evidencia_texto = "Sí" if evidencia is True else "No" if evidencia is False else ""
    agregar_seccion(story, "5. Meta y evidencia", [
        ("Meta esperada", datos.get("meta_esperada")),
        ("¿Se cuenta con evidencia?", evidencia_texto),
        ("Cantidad de archivos fotográficos adjuntos", str(len(fotos)) if fotos else ""),
    ])

    agregar_seccion(story, "6. Valoración técnica y acuerdos", [
        ("Sugerencias y posibilidades de mejora", datos.get("sugerencias")),
        ("Principales acuerdos", datos.get("acuerdos")),
        ("Fecha de la próxima visita de seguimiento", texto_fecha(datos.get("proxima_visita"))),
    ])

    conclusion_elementos = []
    if datos.get("sugerencias"):
        conclusion_elementos.append("se identificaron oportunidades de mejora")
    if datos.get("acuerdos"):
        conclusion_elementos.append("se establecieron acuerdos para su seguimiento")
    if fotos:
        conclusion_elementos.append("se incorporó evidencia fotográfica")
    if conclusion_elementos:
        story.append(Paragraph("7. Conclusión", styles["Seccion"]))
        story.append(Paragraph(
            "La visita permitió documentar la información disponible; " + ", ".join(conclusion_elementos) +
            ". Este informe sirve como respaldo institucional para orientar las acciones posteriores y dar seguimiento a los compromisos registrados.",
            styles["Texto"],
        ))

    if fotos:
        story.append(PageBreak())
        story.append(Paragraph("ANEXO FOTOGRÁFICO", styles["TituloInst"]))
        story.append(Paragraph("Registro visual aportado como evidencia de la visita.", styles["SubtituloInst"]))
        for idx, foto in enumerate(fotos, start=1):
            story.append(KeepTogether([photo_flowable(foto), Paragraph(f"Evidencia fotográfica {idx}", styles["FotoCaption"])]))
            if idx < len(fotos):
                story.append(Spacer(1, 0.25 * cm))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def css_app() -> None:
    st.markdown(
        f"""
        <style>
            .stApp {{ background: #F5F8FB; }}
            .block-container {{ max-width: 1120px; padding-top: 1.2rem; padding-bottom: 3rem; }}
            .hero {{
                background: linear-gradient(135deg, {PRIMARY_BLUE}, {SECONDARY_BLUE});
                color: white; padding: 1.35rem 1.6rem; border-radius: 16px; margin-bottom: 1rem;
                box-shadow: 0 8px 28px rgba(23,59,103,.16);
            }}
            .hero h1 {{ margin: 0; font-size: 2rem; line-height: 1.15; }}
            .hero p {{ margin: .55rem 0 0; opacity: .92; }}
            [data-testid="stForm"] {{
                background: white; padding: 1.2rem 1.25rem; border-radius: 14px;
                border: 1px solid #D9E2EC; box-shadow: 0 4px 16px rgba(15,23,42,.05);
            }}
            h2, h3 {{ color: {PRIMARY_BLUE}; }}
            div[data-testid="stFileUploader"] {{ border-radius: 12px; }}
            .small-note {{ color: #667085; font-size: .9rem; }}
            .required-note {{ color: #B42318; font-size: .86rem; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def obtener_coordenadas_gps() -> Optional[Tuple[float, float]]:
    try:
        location = get_geolocation()
        if location and isinstance(location, dict):
            coords = location.get("coords", {})
            lat = coords.get("latitude")
            lon = coords.get("longitude")
            if lat is not None and lon is not None:
                return float(lat), float(lon)
    except Exception:
        return None
    return None


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")
    css_app()

    st.markdown(
        """
        <div class="hero">
            <h1>Generador de Informes Institucionales</h1>
            <p>Registro de visitas, valoración de actividades preventivas, georreferenciación, evidencia fotográfica y generación automática de PDF.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not Path(DATA_FILE).exists():
        st.error(f"No se encontró el archivo **{DATA_FILE}**. Súbalo al mismo nivel de app.py en el repositorio.")
        st.stop()

    try:
        catalogos = cargar_catalogos(DATA_FILE)
    except Exception as exc:
        st.error(f"No fue posible leer el archivo de datos: {exc}")
        st.stop()

    logos = encontrar_logos()
    if logos:
        st.image(str(logos[0]), width=520)

    if "latitud" not in st.session_state:
        st.session_state.latitud = 9.9281
    if "longitud" not in st.session_state:
        st.session_state.longitud = -84.0907
    if "ubicacion_seleccionada" not in st.session_state:
        st.session_state.ubicacion_seleccionada = False

    with st.container(border=True):
        st.subheader("1. Datos generales de la visita")
        c1, c2 = st.columns(2)
        regiones = catalogos.regiones_delegaciones["Dirección Regional"].dropna().unique().tolist()
        if "DPPP" not in regiones:
            regiones.append("DPPP")
        with c1:
            direccion_regional = st.selectbox(
                "Dirección Regional o dependencia que realiza la visita",
                regiones,
                index=None,
                placeholder="Seleccione una Dirección Regional o DPPP",
                key="direccion_regional",
            )
            delegaciones: List[str] = []
            if direccion_regional and direccion_regional != "DPPP":
                delegaciones = catalogos.regiones_delegaciones.loc[
                    catalogos.regiones_delegaciones["Dirección Regional"] == direccion_regional,
                    "Delegación",
                ].dropna().unique().tolist()
            delegacion_visitada = None
            if direccion_regional and direccion_regional != "DPPP":
                delegacion_visitada = st.selectbox(
                    "Delegación Policial visitada",
                    delegaciones,
                    index=None,
                    placeholder="Seleccione una delegación",
                    key="delegacion_visitada",
                )
        with c2:
            modalidad = st.radio("Modalidad de visita", ["Presencial", "Virtual", "Otro"], horizontal=True, index=None)

        c3, c4 = st.columns(2)
        with c3:
            proposito = st.radio(
                "Propósito de la visita",
                ["Verificación", "Asesoría", "Seguimiento", "Apoyo", "Capacitación"],
                horizontal=True,
                index=None,
            )
        with c4:
            fecha_visita = st.date_input("Fecha de la visita", value=None, format="DD/MM/YYYY")
            hora_visita = st.time_input("Hora de la visita", value=None)

        funcionarios_realizan = st.text_area("Nombre de la(s) persona(s) funcionaria(s) que realiza(n) la visita", height=90)
        funcionarios_atienden = st.text_area("Nombre de la(s) persona(s) funcionaria(s) que atiende(n) la visita", height=90)

    with st.container(border=True):
        st.subheader("2. Ubicación territorial")
        provincias = catalogos.territorios["Provincia"].unique().tolist()
        t1, t2, t3 = st.columns(3)
        with t1:
            provincia = st.selectbox("Provincia", provincias, index=None, placeholder="Seleccione una provincia", key="provincia")
        cantones = [] if not provincia else catalogos.territorios.loc[
            catalogos.territorios["Provincia"] == provincia, "Cantón"
        ].dropna().unique().tolist()
        with t2:
            canton = st.selectbox("Cantón", cantones, index=None, placeholder="Seleccione un cantón", disabled=not provincia, key="canton")
        distritos = [] if not canton else catalogos.territorios.loc[
            (catalogos.territorios["Provincia"] == provincia) & (catalogos.territorios["Cantón"] == canton), "Distritos"
        ].dropna().unique().tolist()
        with t3:
            distrito = st.selectbox("Distrito", distritos, index=None, placeholder="Seleccione un distrito", disabled=not canton, key="distrito")

        top20_auto = es_top20_automatico(catalogos, provincia or "", canton or "", distrito or "")
        if top20_auto is None:
            top20_opcion = st.radio("¿El distrito corresponde al Top 20?", ["Sí", "No"], horizontal=True, index=None)
            es_top20 = True if top20_opcion == "Sí" else False if top20_opcion == "No" else None
        else:
            es_top20 = top20_auto
            st.info(f"Distrito Top 20: **{'Sí' if es_top20 else 'No'}**")

        referencia_lugar = st.text_input(
            "Referencia adicional del lugar",
            placeholder="Ejemplo: oficina regional, centro educativo, salón comunal, comercio u otro punto de referencia",
        )

    with st.container(border=True):
        st.subheader("3. Programa, actividad y acciones relacionadas")
        p1, p2 = st.columns([1, 2])
        with p1:
            programa = st.selectbox(
                "Programa Policial Preventivo y/o actividad",
                PROGRAMAS,
                index=None,
                placeholder="Seleccione una opción",
                key="programa",
            )

        act_df = catalogos.actividades[
            catalogos.actividades["Programa"].str.upper() == (programa or "").upper()
        ] if programa else pd.DataFrame()

        if programa in ACTIVIDADES_POR_PROGRAMA:
            actividades = list(ACTIVIDADES_POR_PROGRAMA[programa])
        else:
            actividades = act_df["Actividad Realizada"].dropna().unique().tolist() if not act_df.empty else []
            if programa and not any(normalizar_texto(x).casefold() == "otros" for x in actividades):
                actividades.append("Otros")

        with p2:
            actividad = st.selectbox(
                "Actividad evaluada o valorada",
                actividades,
                index=None,
                placeholder="Seleccione una actividad",
                disabled=not programa,
                key="actividad",
            )

        responde_a = ""
        if actividad and not act_df.empty and "Responde a:" in act_df.columns:
            coincidencias = act_df.loc[
                act_df["Actividad Realizada"].str.casefold() == actividad.casefold(), "Responde a:"
            ].dropna().tolist()
            responde_a = coincidencias[0] if coincidencias else ""
        if responde_a:
            st.text_area("Marco de planificación", value=responde_a, height=80, disabled=True)

        lineas_accion = st.text_area("Nombres de las acciones, líneas o actividades", height=110)

    with st.container(border=True):
        st.subheader("4. Meta y evidencia")
        meta_esperada = st.text_input("Meta esperada (opcional)", placeholder="Ejemplo: 12 actividades, 150 personas o 4 centros educativos")
        evidencia_opcion = st.radio("¿Se tiene evidencia de la visita o actividad?", ["Sí", "No"], horizontal=True, index=None)
        tiene_evidencia = True if evidencia_opcion == "Sí" else False if evidencia_opcion == "No" else None
        fotos_subidas = st.file_uploader(
            "Suba una o varias fotografías como prueba visual",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            help="Puede adjuntar varias fotografías. Las imágenes se comprimen automáticamente para el PDF.",
        )

    with st.container(border=True):
        st.subheader("5. Valoración, acuerdos y seguimiento")
        sugerencias = st.text_area("Sugerencias y/o posibilidades de mejora", height=120)
        acuerdos = st.text_area("Principales acuerdos", height=120)
        proxima_visita = st.date_input("Fecha de la próxima visita de seguimiento", value=None, format="DD/MM/YYYY")

    st.subheader("6. Georreferenciación")
    st.caption("Puede utilizar el GPS del dispositivo o marcar manualmente el punto exacto en el mapa.")
    g1, g2 = st.columns([1, 2])
    with g1:
        if st.button("Usar GPS del dispositivo", use_container_width=True):
            coords = obtener_coordenadas_gps()
            if coords:
                st.session_state.latitud, st.session_state.longitud = coords
                st.session_state.ubicacion_seleccionada = True
                st.success("Ubicación obtenida correctamente.")
            else:
                st.warning("No fue posible obtener el GPS. Autorice el acceso a la ubicación o marque el punto en el mapa.")
        incluir_coordenadas = st.checkbox("Incluir georreferenciación en el informe", value=st.session_state.ubicacion_seleccionada)
        st.session_state.latitud = st.number_input("Latitud", value=float(st.session_state.latitud), format="%.6f")
        st.session_state.longitud = st.number_input("Longitud", value=float(st.session_state.longitud), format="%.6f")

    with g2:
        mapa = folium.Map(location=[st.session_state.latitud, st.session_state.longitud], zoom_start=14, control_scale=True)
        folium.Marker([st.session_state.latitud, st.session_state.longitud], tooltip="Ubicación seleccionada",
                      icon=folium.Icon(color="blue", icon="info-sign")).add_to(mapa)
        resultado_mapa = st_folium(mapa, height=420, use_container_width=True, returned_objects=["last_clicked"])
        if resultado_mapa and resultado_mapa.get("last_clicked"):
            nuevo = resultado_mapa["last_clicked"]
            st.session_state.latitud = float(nuevo["lat"])
            st.session_state.longitud = float(nuevo["lng"])
            st.session_state.ubicacion_seleccionada = True
            st.info("Punto actualizado. Active la opción de incluir georreferenciación para incorporarlo al informe.")

    enviar = st.button("Preparar informe", type="primary", use_container_width=True)

    if enviar:
        fotos_bytes: List[bytes] = []
        for archivo in fotos_subidas or []:
            try:
                fotos_bytes.append(imagen_a_jpeg_bytes(archivo))
            except Exception as exc:
                st.warning(f"No se pudo procesar la imagen {archivo.name}: {exc}")

        datos = {
            "direccion_regional": direccion_regional or "",
            "delegacion_visitada": delegacion_visitada or "",
            "modalidad": modalidad or "",
            "proposito": proposito or "",
            "fecha_visita": fecha_visita,
            "hora_visita": hora_visita,
            "funcionarios_realizan": funcionarios_realizan,
            "funcionarios_atienden": funcionarios_atienden,
            "provincia": provincia or "",
            "canton": canton or "",
            "distrito": distrito or "",
            "es_top20": es_top20,
            "referencia_lugar": referencia_lugar,
            "programa": programa or "",
            "actividad": actividad or "",
            "responde_a": responde_a,
            "lineas_accion": lineas_accion,
            "meta_esperada": meta_esperada,
            "tiene_evidencia": tiene_evidencia,
            "sugerencias": sugerencias,
            "acuerdos": acuerdos,
            "proxima_visita": proxima_visita,
            "latitud": st.session_state.latitud if incluir_coordenadas else None,
            "longitud": st.session_state.longitud if incluir_coordenadas else None,
        }

        try:
            pdf_bytes = construir_pdf(datos, fotos_bytes, logos)
            st.success("El informe institucional fue generado correctamente.")
            base_nombre = delegacion_visitada or direccion_regional or "Institucional"
            nombre_seguro = re.sub(r"[^A-Za-z0-9_-]+", "_", base_nombre).strip("_") or "Institucional"
            fecha_nombre = fecha_visita.strftime("%Y%m%d") if fecha_visita else datetime.now().strftime("%Y%m%d")
            nombre_pdf = f"Informe_Visita_{nombre_seguro}_{fecha_nombre}.pdf"
            st.download_button("Descargar informe institucional en PDF", data=pdf_bytes, file_name=nombre_pdf,
                               mime="application/pdf", type="primary", use_container_width=True)
            csv_datos = {k: v for k, v in datos.items() if v not in (None, "")}
            st.download_button("Descargar respaldo de datos en CSV",
                               data=pd.DataFrame([csv_datos]).to_csv(index=False).encode("utf-8-sig"),
                               file_name=nombre_pdf.replace(".pdf", ".csv"), mime="text/csv", use_container_width=True)
        except Exception as exc:
            st.exception(exc)


if __name__ == "__main__":
    main()
