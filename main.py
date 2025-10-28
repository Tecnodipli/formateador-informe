import os
import io
import re
import logging
from io import BytesIO
from typing import Optional
from datetime import datetime, timedelta
import secrets

import requests
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openai import OpenAI
from PIL import Image

# =========================
# Configuración de logs
# =========================
logging.basicConfig(level=logging.INFO, format="%(asctime)s — %(levelname)s — %(message)s")
logger = logging.getLogger(__name__)

REPORT_COLOR   = RGBColor(133, 78, 197)
HEADING_COLOR  = RGBColor(85, 54, 185)

# =========================
# Assets predeterminados
# =========================
ASSETS_DIR = "assets"
DEFAULT_PORTADA_PATH       = os.path.join(ASSETS_DIR, "portada.png")
DEFAULT_CONTRAPORTADA_PATH = os.path.join(ASSETS_DIR, "contraportada.png")
DEFAULT_LOGO_PATH          = os.path.join(ASSETS_DIR, "logo.png")

app = FastAPI(title="Formateador de informes")

# =========================
# CORS
# =========================
ALLOWED_ORIGINS = [
    "https://www.dipli.ai",
    "https://dipli.ai",
    "https://isagarcivill09.wixsite.com/turop",
    "https://isagarcivill09.wixsite.com/turop/tienda",
    "https://isagarcivill09-wixsite-com.filesusr.com",
    "https://www.dipli.ai/preparaci%C3%B3n",
    "https://www-dipli-ai.filesusr.com",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    allow_origin_regex=r"https://.*\.filesusr\.com",
)

# =========================
# Descargas temporales en memoria
# =========================
DOWNLOADS: dict[str, tuple[bytes, str, str, datetime]] = {}
DOWNLOAD_TTL_SECS = 900  # 15 min
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

def cleanup_downloads() -> None:
    now = datetime.utcnow()
    expired = [t for t, (_, _, _, exp) in DOWNLOADS.items() if exp <= now]
    for t in expired:
        DOWNLOADS.pop(t, None)

def register_download(data: bytes, filename: str, media_type: str) -> str:
    cleanup_downloads()
    token = secrets.token_urlsafe(16)
    expires_at = datetime.utcnow() + timedelta(seconds=DOWNLOAD_TTL_SECS)
    DOWNLOADS[token] = (data, filename, media_type, expires_at)
    return token

def ensure_default_assets() -> None:
    try:
        os.makedirs(ASSETS_DIR, exist_ok=True)
        if not os.path.exists(DEFAULT_PORTADA_PATH):
            Image.new("RGB", (1200, 1600), (133, 78, 197)).save(DEFAULT_PORTADA_PATH, format="PNG")
        if not os.path.exists(DEFAULT_CONTRAPORTADA_PATH):
            Image.new("RGB", (1200, 1600), (85, 54, 185)).save(DEFAULT_CONTRAPORTADA_PATH, format="PNG")
        if not os.path.exists(DEFAULT_LOGO_PATH):
            img = Image.new("RGB", (600, 600), (255, 255, 255))
            try:
                from PIL import ImageDraw
                draw = ImageDraw.Draw(img)
                draw.ellipse((100, 100, 500, 500), fill=(133, 78, 197))
            except Exception:
                pass
            img.save(DEFAULT_LOGO_PATH, format="PNG")
    except Exception as e:
        logger.warning(f"No se pudieron preparar assets por defecto: {e}")

ensure_default_assets()

# =========================
# Configuración GPT
# =========================
USE_TIKTOKEN = False
MODEL_MAX_TOKENS = 8192

try:
    import tiktoken
    ENCODING = tiktoken.encoding_for_model("gpt-4")
    USE_TIKTOKEN = True
    logger.info("tiktoken disponible.")
except Exception:
    logger.warning("tiktoken no disponible: recorte aproximado por caracteres.")

def trim_to_fit(text: str, reserved_output: int = 700) -> str:
    if USE_TIKTOKEN:
        tokens = ENCODING.encode(text)
        max_input = max(MODEL_MAX_TOKENS - reserved_output - 100, 0)
        return ENCODING.decode(tokens[: max_input if max_input > 0 else 0])
    approx_chars_per_token = 4
    max_input_tokens = max(MODEL_MAX_TOKENS - reserved_output - 100, 0)
    return text[: max_input_tokens * approx_chars_per_token]

def call_gpt(api_key: str, prompt: str, user_input: str, max_tokens: int = 700) -> str:
    if not api_key:
        raise ValueError("API Key de OpenAI es requerida")
    client = OpenAI(api_key=api_key)
    trimmed_input = trim_to_fit(user_input, reserved_output=max_tokens)

    def try_model(model_name: str) -> str:
        resp = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": trimmed_input},
            ],
            temperature=0.5,
            max_tokens=max_tokens,
        )
        return resp.choices[0].message.content.strip()

    try:
        return try_model("gpt-4o-mini")
    except Exception as e:
        msg = str(e)
        logger.error(f"Error GPT: {msg}")
        if any(x in msg for x in ("insufficient_quota", "429", "Rate limit")):
            try:
                return try_model("gpt-3.5-turbo")
            except Exception as fallback_err:
                logger.error("Fallback también falló: %s", fallback_err)
                return "Contenido no disponible."
        return "Contenido no disponible."

# =========================
# Prompts
# =========================
PROMPT_RESUMEN = """
Actúa como un experto en redacción ejecutiva y análisis de informes cualitativos. Tu tarea es redactar un resumen ejecutivo profesional y conciso, basándote en el siguiente contenido del informe.

**Requisitos:**
* Longitud: proporcional al documento original (~3% del total de palabras).
* Estructura: dos párrafos integrados.
* Contenido: objetivo, alcance, hallazgos, impacto y recomendaciones.
* Tono: formal, claro y accesible.
* Evitar: repeticiones, tecnicismos y explicaciones extensas.
"""

PROMPT_HALLAZGOS = """
Actúa como un experto en redacción ejecutiva y análisis cualitativo. A partir del siguiente documento, redacta una sección titulada "Principales Hallazgos" en español.

**Requisitos:**
* Enfócate en los hallazgos clave y sus implicaciones.
* Estructura: numeración y contexto breve por punto.
* Estilo: profesional, claro, útil para tomadores de decisiones.
* Evitar: recomendaciones, juicios, generalidades o conclusiones.
"""

# =========================
# Formateo
# =========================
def modify_style(doc: Document, style_name: str, size_pt: int,
                 bold: bool = False, italic: bool = False,
                 color: Optional[RGBColor] = None) -> None:
    font = doc.styles[style_name].font
    font.name = "Century Gothic"
    font.size = Pt(size_pt)
    font.bold = bold
    font.italic = italic
    if color:
        font.color.rgb = color

def insert_cover_page(doc_out: Document, portada_path) -> None:
    sec0 = doc_out.sections[0]
    orig = {
        'top': sec0.top_margin,
        'bottom': sec0.bottom_margin,
        'left': sec0.left_margin,
        'right': sec0.right_margin,
    }
    sec0.top_margin = sec0.bottom_margin = Cm(0)
    sec0.left_margin = sec0.right_margin = Cm(0)

    pw, ph = sec0.page_width, sec0.page_height
    try:
        if isinstance(portada_path, list):
            img_url = portada_path[0]
            resp = requests.get(img_url, timeout=15)
            resp.raise_for_status()
            img = BytesIO(resp.content)
        else:
            img = portada_path
    except Exception as e:
        logger.error(f"Error cargando imagen de portada: {e}")
        return

    p = doc_out.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run()
    run.add_picture(img, width=pw, height=ph)

    doc_out.add_page_break()
    new_sec = doc_out.add_section(WD_SECTION.NEW_PAGE)
    for k, v in orig.items():
        setattr(new_sec, f"{k}_margin", v)

def insert_contraportada_body(doc_out: Document, contraportada_path) -> None:
    last_sec = doc_out.sections[-1]
    orig = {
        'top': last_sec.top_margin,
        'bottom': last_sec.bottom_margin,
        'left': last_sec.left_margin,
        'right': last_sec.right_margin,
    }
    cover_sec = doc_out.add_section(WD_SECTION.NEW_PAGE)
    cover_sec.top_margin = cover_sec.bottom_margin = Cm(0)
    cover_sec.left_margin = cover_sec.right_margin = Cm(0)

    pw, ph = cover_sec.page_width, cover_sec.page_height
    try:
        if isinstance(contraportada_path, str):
            resp = requests.get(contraportada_path, timeout=15)
            resp.raise_for_status()
            img = BytesIO(resp.content)
        else:
            img = contraportada_path
    except Exception as e:
        logger.error(f"Error descargando contraportada: {e}")
        return

    p = doc_out.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run()
    run.add_picture(img, width=pw, height=ph)

    new_sec = doc_out.add_section(WD_SECTION.NEW_PAGE)
    for k, v in orig.items():
        setattr(new_sec, f"{k}_margin", v)

def insert_footer_logo(doc: Document, logo_source) -> None:
    sec = doc.sections[-1]
    footer = sec.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    try:
        if isinstance(logo_source, str):
            resp = requests.get(logo_source, timeout=15)
            resp.raise_for_status()
            img = BytesIO(resp.content)
        else:
            img = logo_source
        p.add_run().add_picture(img, width=Inches(1.5))
    except Exception as e:
        logger.error(f"Error insertando logo en footer: {e}")

def add_table_of_contents(paragraph) -> None:
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'TOC \\o "1-5" \\h \\z \\u')
    paragraph._p.append(fld)

def is_title3(prev: str, curr: str, next_: str) -> bool:
    return (
        prev.strip() == "" and next_.strip() == "" and
        curr.startswith("**") and curr.endswith("**") and
        ":" not in curr
    )

def is_title5(line: str) -> bool:
    return bool(re.match(r"^\s*(\d+\.\s+|\-\s+)\*\*.*?\*\*:", line.strip()))

def extract_title5_text(line: str) -> str:
    m = re.search(r"\*\*(.*?)\*\*", line)
    return m.group(1).strip(": ") if m else ""

# ========= NUEVO: verbatim centrado con saltos =========
def format_text_block(doc: Document, texto: str, color=RGBColor(133, 78, 197)) -> None:
    tokens = re.split(r'(\*\*[^*]+\*\*|_[^_]+_|[*][^*]+[*]|_"[^"]+"_)', texto)
    p = None

    def ensure_p():
        nonlocal p
        if p is None:
            p = doc.add_paragraph("", style="Normal")
            p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        return p

    for token in tokens:
        if not token:
            continue
        if token.startswith('_"') and token.endswith('"_'):
            doc.add_paragraph("")  # línea en blanco antes
            vp = doc.add_paragraph("", style="Normal")
            vp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = vp.add_run(token[2:-2])
            run.bold = True
            run.italic = True
            run.font.color.rgb = color
            doc.add_paragraph("")  # línea en blanco después
            p = None
            continue
        if token.startswith("**") and token.endswith("**"):
            run = ensure_p().add_run(token[2:-2])
            run.bold = True
            continue
        if token.startswith("*") and token.endswith("*"):
            run = ensure_p().add_run(token[1:-1])
            run.bold = True
            run.italic = True
            run.font.color.rgb = color
            continue
        if token.startswith("_") and token.endswith("_"):
            run = ensure_p().add_run(token[1:-1])
            run.italic = True
            continue
        ensure_p().add_run(token)

# ====== Helpers para evitar duplicados de títulos ======
_NBSP = "\u00A0"
_ZWSP = "\u200B"

def _normalize_ws(s: str) -> str:
    # Reemplaza NBSP y ZWSP por espacio normal y colapsa espacios
    s = s.replace(_NBSP, " ").replace(_ZWSP, "")
    return re.sub(r"\s+", " ", s, flags=re.UNICODE).strip()

def _normalize_title(s: str) -> str:
    s = _normalize_ws(s)
    s = re.sub(r"^[#\s]+", "", s, flags=re.UNICODE)  # quita # y espacios
    s = s.strip("*_ :\t-—")  # limpieza ligera
    return s.lower()

def _is_duplicate_section_title(texto: str) -> bool:
    norm = _normalize_title(texto)
    return norm.startswith("resumen ejecutivo") or norm.startswith("principales hallazgos")

def __md_heading_info(s: str):
    """
    Detecta encabezados Markdown con cualquier espacio unicode.
    Devuelve (level:int|None, title:str|None).
    """
    s2 = _normalize_ws(s)
    m = re.match(r"^(#{1,6})\s*(.+?)\s*$", s2, flags=re.UNICODE)
    if not m:
        return None, None
    level = len(m.group(1))
    title = m.group(2)
    return level, title

def _strip_duplicate_heading_lines(text: str) -> str:
    # Elimina líneas que sean exactamente el mismo título (con o sin ###)
    lines = text.splitlines()
    out = []
    for ln in lines:
        lvl, ttl = __md_heading_info(ln)
        candidate = ttl if ttl else ln
        if _is_duplicate_section_title(candidate):
            continue
        out.append(ln)
    return "\n".join(out)

# =========================
# Generación del informe
# =========================
def generate_report(api_key: str,
                    input_doc_bytes: bytes,
                    portada_bytes: Optional[bytes],
                    contraportada_bytes: Optional[bytes],
                    logo_bytes: Optional[bytes],
                    use_defaults: bool,
                    filename_hint: str) -> bytes:
    doc = Document(BytesIO(input_doc_bytes))
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    full_text  = "\n\n".join(paragraphs)

    if use_defaults:
        with open(DEFAULT_PORTADA_PATH, "rb") as f:
            portada_path = BytesIO(f.read())
        with open(DEFAULT_CONTRAPORTADA_PATH, "rb") as f:
            contraportada_path = BytesIO(f.read())
        with open(DEFAULT_LOGO_PATH, "rb") as f:
            logo_path = BytesIO(f.read())
    else:
        if not (portada_bytes and contraportada_bytes and logo_bytes):
            raise ValueError("Faltan imágenes personalizadas (portada/contraportada/logo).")
        portada_path       = BytesIO(portada_bytes)
        contraportada_path = BytesIO(contraportada_bytes)
        logo_path          = BytesIO(logo_bytes)

    doc_out = Document()
    insert_cover_page(doc_out, portada_path)

    modify_style(doc_out, 'Normal',    12)
    modify_style(doc_out, 'Heading 1', 20, bold=True,  color=HEADING_COLOR)
    modify_style(doc_out, 'Heading 2', 16, bold=True,  color=HEADING_COLOR)
    modify_style(doc_out, 'Heading 3', 14, bold=True,  color=HEADING_COLOR)
    modify_style(doc_out, 'Heading 4', 12, bold=True,  color=HEADING_COLOR)
    modify_style(doc_out, 'Heading 5', 20, bold=True,  color=REPORT_COLOR)

    # TOC
    doc_out.add_paragraph("Tabla de contenidos", style="Heading 5")
    add_table_of_contents(doc_out.add_paragraph())

    # Resumen Ejecutivo
    resumen = call_gpt(api_key, PROMPT_RESUMEN, full_text[:10000], 500)
    doc_out.add_paragraph("Resumen Ejecutivo", style="Heading 1")
    doc_out.add_paragraph(resumen, style="Normal")

    # Principales Hallazgos
    hallazgos_text = full_text[:10000]
    hallazgos = call_gpt(api_key, PROMPT_HALLAZGOS, hallazgos_text, 700)
    # Limpia encabezados duplicados dentro del texto generado
    hallazgos = _strip_duplicate_heading_lines(hallazgos)
    doc_out.add_paragraph("Principales Hallazgos", style="Heading 1")
    for item in hallazgos.split("\n"):
        t = item.strip()
        if not t:
            continue
        format_text_block(doc_out, t, color=REPORT_COLOR)

    # Cuerpo formateado (evita duplicar títulos ya agregados)
    for i, para in enumerate(paragraphs):
        raw = para
        t = raw.strip()
        if not t:
            continue

        # 1) Si la línea (con o sin ###/negritas) normalizada coincide con título duplicado, se omite
        if _is_duplicate_section_title(t):
            continue

        # 2) Si es un encabezado Markdown, manejarlo robustamente
        lvl, ttl = __md_heading_info(t)
        if lvl is not None:
            if _is_duplicate_section_title(ttl or ""):
                continue
            if lvl == 3:  # equivalente al antiguo "### "
                insert_contraportada_body(doc_out, contraportada_path)
                doc_out.add_paragraph(ttl, style="Heading 1")
            elif lvl == 4:
                doc_out.add_paragraph(ttl, style="Heading 2")
            else:
                # Otros niveles: mapear de forma suave
                style = "Heading 3" if lvl == 5 else "Heading 4"
                doc_out.add_paragraph(ttl, style=style)
            continue

        prev = paragraphs[i-1] if i > 0 else ""
        nxt  = paragraphs[i+1] if i < len(paragraphs)-1 else ""

        if is_title3(prev, t, nxt):
            posible_titulo = t.strip("*").strip()
            if _is_duplicate_section_title(posible_titulo):
                continue
            doc_out.add_paragraph(posible_titulo, style="Heading 3")
            continue

        if is_title5(t):
            posible_titulo = extract_title5_text(t)
            if _is_duplicate_section_title(posible_titulo):
                continue
            doc_out.add_paragraph(posible_titulo, style="Heading 4")
            continue

        # Párrafos normales + verbatims
        format_text_block(doc_out, t, color=REPORT_COLOR)

    insert_footer_logo(doc_out, logo_path)

    out_bytes = io.BytesIO()
    doc_out.save(out_bytes)
    out_bytes.seek(0)
    return out_bytes.getvalue()

# ---------- Endpoints ----------
@app.post("/generate-report-link")
async def generate_report_link(
    request: Request,
    file: UploadFile = File(..., description="Archivo .docx base"),
    openai_api_key: str = Form(..., description="Tu OpenAI API Key"),
    usar_personalizadas: bool = Form(False),
    portada: UploadFile | None = File(None),
    contraportada: UploadFile | None = File(None),
    logo: UploadFile | None = File(None),
):
    if not openai_api_key:
        raise HTTPException(status_code=400, detail="openai_api_key es requerida")
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Debes subir un archivo .docx válido.")

    base_bytes = await file.read()
    portada_bytes = await portada.read() if (portada and usar_personalizadas) else None
    contraportada_bytes = await contraportada.read() if (contraportada and usar_personalizadas) else None
    logo_bytes = await logo.read() if (logo and usar_personalizadas) else None

    result_bytes = generate_report(
        api_key=openai_api_key,
        input_doc_bytes=base_bytes,
        portada_bytes=portada_bytes,
        contraportada_bytes=contraportada_bytes,
        logo_bytes=logo_bytes,
        use_defaults=not usar_personalizadas,
        filename_hint=file.filename,
    )

    final_name = file.filename.replace(".docx", "") + "_INFORME_FINAL.docx"
    token = register_download(result_bytes, final_name, DOCX_MEDIA_TYPE)

    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/download/{token}"

    return {"download_url": download_url, "expires_in_seconds": DOWNLOAD_TTL_SECS}

@app.post("/generate-report-simple-link")
async def generate_report_simple_link(
    request: Request,
    file: UploadFile = File(..., description="Archivo .docx base"),
    openai_api_key: str = Form(..., description="Tu OpenAI API Key"),
):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Debes subir un archivo .docx válido.")

    base_bytes = await file.read()
    result_bytes = generate_report(
        api_key=openai_api_key,
        input_doc_bytes=base_bytes,
        portada_bytes=None,
        contraportada_bytes=None,
        logo_bytes=None,
        use_defaults=True,
        filename_hint=file.filename,
    )

    final_name = file.filename.replace(".docx", "") + "_INFORME_FINAL.docx"
    token = register_download(result_bytes, final_name, DOCX_MEDIA_TYPE)

    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/download/{token}"

    return {"download_url": download_url, "expires_in_seconds": DOWNLOAD_TTL_SECS}

@app.get("/download/{token}")
def download_token(token: str):
    cleanup_downloads()
    item = DOWNLOADS.get(token)
    if not item:
        raise HTTPException(status_code=404, detail="Link expirado o inválido")
    data, filename, media_type, exp = item
    if exp <= datetime.utcnow():
        DOWNLOADS.pop(token, None)
        raise HTTPException(status_code=410, detail="Link expirado")
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Cache-Control": "no-store",
    }
    return StreamingResponse(io.BytesIO(data), media_type=media_type, headers=headers)

@app.get("/")
async def root():
    return {"message": "API de Generador de Informes DOCX funcionando", "version": "1.0.0"}

@app.get("/health")
async def health_check():
    return {"status": "healthy", "message": "API funcionando correctamente"}

