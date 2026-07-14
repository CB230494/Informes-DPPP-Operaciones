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
PROGRAMAS = ["GREAT", "DARE", "VIFA", "MPAS", "PSCC"]
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
    font_regular, font_bold = registrar_fuente()
    buffer = io.BytesIO()

    doc = BaseDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=1.65 * cm,
        leftMargin=1.65 * cm,
        topMargin=2.35 * cm,
        bottomMargin=1.45 * cm,
        title=f"Informe institucional - {datos['delegacion_visitada']}",
        author="Ministerio de Seguridad Pública",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="contenido")
    template = PageTemplate(
        id="institucional",
        frames=[frame],
        onPage=lambda canvas, d: encabezado_pie(canvas, d, logos, font_regular, font_bold),
    )
    doc.addPageTemplates([template])

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TituloInst",
        parent=styles["Title"],
        fontName=font_bold,
        fontSize=16,
        leading=20,
        textColor=colors.HexColor(PRIMARY_BLUE),
        alignment=TA_CENTER,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name="SubtituloInst",
        parent=styles["Normal"],
        fontName=font_regular,
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#475467"),
        alignment=TA_CENTER,
        spaceAfter=14,
    ))
    styles.add(ParagraphStyle(
        name="Seccion",
        parent=styles["Heading2"],
        fontName=font_bold,
        fontSize=11.5,
        leading=14,
        textColor=colors.white,
        backColor=colors.HexColor(PRIMARY_BLUE),
        borderPadding=(5, 7, 5, 7),
        spaceBefore=9,
        spaceAfter=7,
    ))
    styles.add(ParagraphStyle(
        name="Texto",
        parent=styles["BodyText"],
        fontName=font_regular,
        fontSize=9.1,
        leading=13,
        textColor=colors.HexColor(DARK_TEXT),
        alignment=TA_JUSTIFY,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="CeldaEtiqueta",
        parent=styles["BodyText"],
        fontName=font_bold,
        fontSize=8.2,
        leading=10.5,
        textColor=colors.HexColor(PRIMARY_BLUE),
    ))
    styles.add(ParagraphStyle(
        name="CeldaValor",
        parent=styles["BodyText"],
        fontName=font_regular,
        fontSize=8.2,
        leading=10.5,
        textColor=colors.HexColor(DARK_TEXT),
    ))
    styles.add(ParagraphStyle(
        name="FotoCaption",
        parent=styles["BodyText"],
        fontName=font_regular,
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#667085"),
        alignment=TA_CENTER,
        spaceAfter=8,
    ))

    story: List[Any] = []
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph("INFORME INSTITUCIONAL DE VISITA", styles["TituloInst"]))
    story.append(Paragraph(
        f"{datos['direccion_regional']} - {datos['delegacion_visitada']}<br/>"
        f"Fecha de la visita: {datos['fecha_visita'].strftime('%d/%m/%Y')} a las {datos['hora_visita'].strftime('%H:%M')}",
        styles["SubtituloInst"],
    ))

    intro = (
        "El presente informe documenta la visita institucional realizada con el propósito de "
        f"<b>{datos['proposito'].lower()}</b>, en el marco del seguimiento técnico y operativo de los "
        "Programas Policiales Preventivos. El documento consolida la información territorial, las personas "
        "participantes, la actividad valorada, el nivel de cumplimiento de la meta, la evidencia disponible, "
        "las oportunidades de mejora y los acuerdos adoptados, con el fin de facilitar la trazabilidad, la "
        "toma de decisiones y el seguimiento de los compromisos establecidos."
    )
    story.append(Paragraph("1. Introducción", styles["Seccion"]))
    story.append(Paragraph(intro, styles["Texto"]))

    def p(text: Any, style="CeldaValor") -> Paragraph:
        safe = normalizar_texto(text) or "No indicado"
        return Paragraph(safe.replace("\n", "<br/>"), styles[style])

    def tabla_datos(filas: List[Tuple[str, Any]], widths=(5.2 * cm, 11.0 * cm)) -> Table:
        data = [[p(etq, "CeldaEtiqueta"), p(val)] for etq, val in filas]
        tabla = Table(data, colWidths=list(widths), repeatRows=0, hAlign="LEFT")
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(LIGHT_BLUE)),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return tabla

    story.append(Paragraph("2. Información general de la visita", styles["Seccion"]))
    story.append(tabla_datos([
        ("Dirección Regional que realiza la visita", datos["direccion_regional"]),
        ("Modalidad", datos["modalidad"]),
        ("Propósito", datos["proposito"]),
        ("Fecha y hora", f"{datos['fecha_visita'].strftime('%d/%m/%Y')} - {datos['hora_visita'].strftime('%H:%M')}"),
        ("Persona(s) funcionaria(s) que realiza(n) la visita", datos["funcionarios_realizan"]),
        ("Persona(s) funcionaria(s) que atiende(n) la visita", datos["funcionarios_atienden"]),
    ]))

    story.append(Paragraph("3. Ubicación y referencia territorial", styles["Seccion"]))
    top20_text = "Sí" if datos["es_top20"] else "No"
    ubicacion_filas = [
        ("Provincia", datos["provincia"]),
        ("Cantón", datos["canton"]),
        ("Distrito", datos["distrito"]),
        ("Distrito perteneciente al Top 20", top20_text),
        ("Delegación policial visitada", datos["delegacion_visitada"]),
        ("Referencia del lugar", datos["referencia_lugar"]),
    ]
    if datos.get("latitud") is not None and datos.get("longitud") is not None:
        mapa_url = f"https://www.openstreetmap.org/?mlat={datos['latitud']}&mlon={datos['longitud']}#map=17/{datos['latitud']}/{datos['longitud']}"
        ubicacion_filas.extend([
            ("Coordenadas", f"Latitud: {datos['latitud']:.6f} | Longitud: {datos['longitud']:.6f}"),
            ("Enlace de ubicación", f'<link href="{mapa_url}" color="#1D4ED8">Abrir ubicación en mapa</link>'),
        ])
    story.append(tabla_datos(ubicacion_filas))

    story.append(Paragraph("4. Programa y actividad valorada", styles["Seccion"]))
    story.append(tabla_datos([
        ("Programa Policial Preventivo", datos["programa"]),
        ("Actividad evaluada o valorada", datos["actividad"]),
        ("Marco de planificación", datos["responde_a"]),
        ("Línea(s) de acción relacionada(s)", datos["lineas_accion"]),
    ]))

    story.append(Paragraph("5. Cumplimiento de la meta y evidencia", styles["Seccion"]))
    avance = numero_seguro(datos["avance_porcentaje"])
    cumplimiento = "Dentro del rango esperado" if avance >= 80 else "Requiere seguimiento" if avance >= 50 else "Requiere atención prioritaria"
    story.append(tabla_datos([
        ("Meta esperada", datos["meta_esperada"]),
        ("Avance en el cumplimiento", f"{avance:.1f}%"),
        ("Valoración general", cumplimiento),
        ("¿Se cuenta con evidencia?", "Sí" if datos["tiene_evidencia"] else "No"),
        ("Cantidad de archivos fotográficos adjuntos", str(len(fotos))),
    ]))

    story.append(Paragraph("6. Valoración técnica y acuerdos", styles["Seccion"]))
    story.append(tabla_datos([
        ("Sugerencias y posibilidades de mejora", datos["sugerencias"]),
        ("Principales acuerdos", datos["acuerdos"]),
        ("Fecha de la próxima visita de seguimiento", datos["proxima_visita"].strftime("%d/%m/%Y") if datos.get("proxima_visita") else "No definida"),
    ]))

    story.append(Paragraph("7. Conclusión", styles["Seccion"]))
    conclusion = (
        "La visita permitió registrar el estado de la actividad y el avance de la meta, identificar los elementos "
        "que requieren fortalecimiento y establecer acuerdos para su seguimiento. La información consignada en "
        "este informe constituye un respaldo institucional para verificar el cumplimiento de los compromisos y "
        "orientar las acciones posteriores de la Dirección Regional, la Delegación Policial y el Programa Policial "
        "Preventivo correspondiente."
    )
    story.append(Paragraph(conclusion, styles["Texto"]))

    if fotos:
        story.append(PageBreak())
        story.append(Paragraph("ANEXO FOTOGRÁFICO", styles["TituloInst"]))
        story.append(Paragraph(
            "Registro visual aportado como evidencia de la visita y de la actividad valorada.",
            styles["SubtituloInst"],
        ))
        for idx, foto in enumerate(fotos, start=1):
            bloque = [
                photo_flowable(foto),
                Paragraph(f"Evidencia fotográfica {idx}", styles["FotoCaption"]),
            ]
            story.append(KeepTogether(bloque))
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
    else:
        st.warning("No se encontró Logo1.jpeg/png/jpg. La aplicación funcionará, pero el PDF se generará sin logotipo.")

    if "latitud" not in st.session_state:
        st.session_state.latitud = 9.9281
    if "longitud" not in st.session_state:
        st.session_state.longitud = -84.0907

    with st.form("formulario_informe", clear_on_submit=False):
        st.subheader("1. Datos generales de la visita")
        c1, c2 = st.columns(2)
        regiones = catalogos.regiones_delegaciones["Dirección Regional"].dropna().unique().tolist()
        with c1:
            direccion_regional = st.selectbox("Dirección Regional que realiza la visita *", regiones, index=None, placeholder="Seleccione una Dirección Regional")
        with c2:
            modalidad = st.radio("Modalidad de visita *", ["Presencial", "Virtual", "Otro"], horizontal=True)

        c3, c4 = st.columns(2)
        with c3:
            proposito = st.radio("Propósito de la visita *", ["Verificación", "Asesoría", "Seguimiento"], horizontal=True)
        with c4:
            fecha_visita = st.date_input("Fecha de la visita *", value=date.today(), format="DD/MM/YYYY")
            hora_visita = st.time_input("Hora de la visita *", value=datetime.now().time().replace(second=0, microsecond=0))

        funcionarios_realizan = st.text_area("Nombre de la(s) persona(s) funcionaria(s) que realiza(n) la visita *", height=90)
        funcionarios_atienden = st.text_area("Nombre de la(s) persona(s) funcionaria(s) que atiende(n) la visita *", height=90)

        st.subheader("2. Ubicación territorial y dependencia visitada")
        provincias = catalogos.territorios["Provincia"].unique().tolist()
        t1, t2, t3 = st.columns(3)
        with t1:
            provincia = st.selectbox("Provincia *", provincias, index=None, placeholder="Seleccione una provincia")
        cantones = [] if not provincia else catalogos.territorios.loc[catalogos.territorios["Provincia"] == provincia, "Cantón"].unique().tolist()
        with t2:
            canton = st.selectbox("Cantón *", cantones, index=None, placeholder="Seleccione un cantón", disabled=not provincia)
        distritos = [] if not canton else catalogos.territorios.loc[
            (catalogos.territorios["Provincia"] == provincia) & (catalogos.territorios["Cantón"] == canton), "Distritos"
        ].unique().tolist()
        with t3:
            distrito = st.selectbox("Distrito *", distritos, index=None, placeholder="Seleccione un distrito", disabled=not canton)

        top20_auto = es_top20_automatico(catalogos, provincia or "", canton or "", distrito or "")
        d1, d2 = st.columns(2)
        with d1:
            if top20_auto is None:
                es_top20 = st.radio("¿El distrito corresponde al Top 20? *", ["No", "Sí"], horizontal=True) == "Sí"
                if not catalogos.top20_col:
                    st.caption("El Excel aún no incluye una columna Top 20; por ahora la clasificación se registra manualmente.")
            else:
                es_top20 = top20_auto
                st.info(f"Clasificación Top 20 detectada en el Excel: **{'Sí' if es_top20 else 'No'}**")
        delegaciones = []
        if direccion_regional:
            delegaciones = catalogos.regiones_delegaciones.loc[
                catalogos.regiones_delegaciones["Dirección Regional"] == direccion_regional, "Delegación"
            ].dropna().unique().tolist()
        with d2:
            delegacion_visitada = st.selectbox("Delegación Policial visitada *", delegaciones, index=None, placeholder="Seleccione una delegación", disabled=not direccion_regional)

        referencia_lugar = st.text_input("Referencia adicional del lugar", placeholder="Ejemplo: oficina regional, centro educativo, salón comunal, comercio u otro punto de referencia")

        st.subheader("3. Programa, actividad y línea de acción")
        p1, p2 = st.columns([1, 2])
        with p1:
            programa = st.selectbox("Programa Policial Preventivo *", PROGRAMAS, index=None, placeholder="Seleccione un programa")
        act_df = catalogos.actividades[catalogos.actividades["Programa"].str.upper() == (programa or "").upper()] if programa else pd.DataFrame()
        actividades = act_df["Actividad Realizada"].dropna().unique().tolist() if not act_df.empty else []
        with p2:
            actividad = st.selectbox("Actividad evaluada o valorada *", actividades, index=None, placeholder="Seleccione una actividad", disabled=not programa)

        responde_a = ""
        if actividad and not act_df.empty and "Responde a:" in act_df.columns:
            coincidencias = act_df.loc[act_df["Actividad Realizada"] == actividad, "Responde a:"].dropna().tolist()
            responde_a = coincidencias[0] if coincidencias else ""
            if responde_a:
                st.text_area("Marco de planificación (tomado del Excel)", value=responde_a, height=80, disabled=True)

        lineas_accion = st.text_area("Nombre de la(s) línea(s) de acción relacionada(s) *", height=100)

        st.subheader("4. Meta, avance y evidencia")
        m1, m2 = st.columns(2)
        with m1:
            meta_esperada = st.text_input("Meta esperada *", placeholder="Ejemplo: 12 actividades, 150 personas, 4 centros educativos")
        with m2:
            avance_porcentaje = st.number_input("Avance en el cumplimiento de la meta (%) *", min_value=0.0, max_value=100.0, value=0.0, step=1.0)

        tiene_evidencia = st.radio("¿Se tiene evidencia del avance de la meta? *", ["Sí", "No"], horizontal=True) == "Sí"
        fotos_subidas = st.file_uploader(
            "Suba una o varias fotografías como prueba visual",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            help="Puede adjuntar varias fotografías. Las imágenes se comprimen automáticamente para el PDF.",
        )

        st.subheader("5. Valoración, acuerdos y seguimiento")
        sugerencias = st.text_area("Sugerencias y/o posibilidades de mejora *", height=120)
        acuerdos = st.text_area("Principales acuerdos *", height=120)
        proxima_visita = st.date_input("Fecha de la próxima visita de seguimiento", value=None, format="DD/MM/YYYY")

        st.markdown('<p class="required-note">Los campos marcados con * son obligatorios.</p>', unsafe_allow_html=True)
        enviar = st.form_submit_button("Validar datos y preparar informe", type="primary", use_container_width=True)

    st.subheader("6. Georreferenciación")
    st.caption("Puede utilizar el GPS del dispositivo o marcar manualmente el punto exacto en el mapa.")
    g1, g2 = st.columns([1, 2])
    with g1:
        if st.button("Usar GPS del dispositivo", use_container_width=True):
            coords = obtener_coordenadas_gps()
            if coords:
                st.session_state.latitud, st.session_state.longitud = coords
                st.success("Ubicación obtenida correctamente.")
            else:
                st.warning("No fue posible obtener el GPS. Autorice el acceso a la ubicación en el navegador o marque el punto en el mapa.")
        st.session_state.latitud = st.number_input("Latitud", value=float(st.session_state.latitud), format="%.6f")
        st.session_state.longitud = st.number_input("Longitud", value=float(st.session_state.longitud), format="%.6f")
        st.markdown('<p class="small-note">Las coordenadas se incorporarán al PDF y se generará un enlace de ubicación.</p>', unsafe_allow_html=True)

    with g2:
        mapa = folium.Map(location=[st.session_state.latitud, st.session_state.longitud], zoom_start=14, control_scale=True)
        folium.Marker(
            [st.session_state.latitud, st.session_state.longitud],
            tooltip="Ubicación seleccionada",
            icon=folium.Icon(color="blue", icon="info-sign"),
        ).add_to(mapa)
        resultado_mapa = st_folium(mapa, height=420, use_container_width=True, returned_objects=["last_clicked"])
        if resultado_mapa and resultado_mapa.get("last_clicked"):
            nuevo = resultado_mapa["last_clicked"]
            st.session_state.latitud = float(nuevo["lat"])
            st.session_state.longitud = float(nuevo["lng"])
            st.info("Punto actualizado. Presione nuevamente “Validar datos y preparar informe” para generar el PDF con esta ubicación.")

    if enviar:
        requeridos = {
            "Dirección Regional": direccion_regional,
            "Provincia": provincia,
            "Cantón": canton,
            "Distrito": distrito,
            "Delegación visitada": delegacion_visitada,
            "Programa": programa,
            "Actividad": actividad,
            "Persona(s) que realizan la visita": funcionarios_realizan,
            "Persona(s) que atienden la visita": funcionarios_atienden,
            "Línea(s) de acción": lineas_accion,
            "Meta esperada": meta_esperada,
            "Sugerencias": sugerencias,
            "Acuerdos": acuerdos,
        }
        faltantes = [campo for campo, valor in requeridos.items() if not normalizar_texto(valor)]
        if faltantes:
            st.error("Complete los siguientes campos obligatorios: " + ", ".join(faltantes))
            st.stop()

        if tiene_evidencia and not fotos_subidas:
            st.warning("Se indicó que existe evidencia, pero no se adjuntaron fotografías. El informe se generará sin anexo fotográfico.")

        fotos_bytes: List[bytes] = []
        for archivo in fotos_subidas or []:
            try:
                fotos_bytes.append(imagen_a_jpeg_bytes(archivo))
            except Exception as exc:
                st.warning(f"No se pudo procesar la imagen {archivo.name}: {exc}")

        datos = {
            "direccion_regional": direccion_regional,
            "modalidad": modalidad,
            "proposito": proposito,
            "fecha_visita": fecha_visita,
            "hora_visita": hora_visita,
            "funcionarios_realizan": funcionarios_realizan,
            "funcionarios_atienden": funcionarios_atienden,
            "provincia": provincia,
            "canton": canton,
            "distrito": distrito,
            "es_top20": es_top20,
            "delegacion_visitada": delegacion_visitada,
            "referencia_lugar": referencia_lugar,
            "programa": programa,
            "actividad": actividad,
            "responde_a": responde_a or "No indicado",
            "lineas_accion": lineas_accion,
            "meta_esperada": meta_esperada,
            "avance_porcentaje": avance_porcentaje,
            "tiene_evidencia": tiene_evidencia,
            "sugerencias": sugerencias,
            "acuerdos": acuerdos,
            "proxima_visita": proxima_visita,
            "latitud": st.session_state.latitud,
            "longitud": st.session_state.longitud,
        }

        try:
            pdf_bytes = construir_pdf(datos, fotos_bytes, logos)
            st.success("El informe institucional fue generado correctamente.")
            nombre_seguro = re.sub(r"[^A-Za-z0-9_-]+", "_", delegacion_visitada).strip("_")
            nombre_pdf = f"Informe_Visita_{nombre_seguro}_{fecha_visita.strftime('%Y%m%d')}.pdf"
            st.download_button(
                "Descargar informe institucional en PDF",
                data=pdf_bytes,
                file_name=nombre_pdf,
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
            st.download_button(
                "Descargar respaldo de datos en CSV",
                data=pd.DataFrame([datos]).to_csv(index=False).encode("utf-8-sig"),
                file_name=nombre_pdf.replace(".pdf", ".csv"),
                mime="text/csv",
                use_container_width=True,
            )
        except Exception as exc:
            st.exception(exc)


if __name__ == "__main__":
    main()
