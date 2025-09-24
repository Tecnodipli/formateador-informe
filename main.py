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
# CORS: habilitar solo tus dominios
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
    allow_origins=ALLOWED_ORIGINS,   # ✅ ahora sí aplica tu lista
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    allow_origin_regex=r"https://.*\.filesusr\.com",
)

# =========================
# Descargas temporales en memoria
# =========================
DOWNLOADS: dict[str, tuple[bytes, str, str, datetime]] = {}
DOWNLOAD_TTL_SECS = 900  # 15 minutos
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

def cleanup_downloads() -> None:
    """Elimina descargas expiradas de memoria"""
    now = datetime.utcnow()
    expired = [t for t, (_, _, _, exp) in DOWNLOADS.items() if exp <= now]
    for t in expired:
        DOWNLOADS.pop(t, None)

def register_download(data: bytes, filename: str, media_type: str) -> str:
    """Guarda un archivo en memoria y devuelve un token único"""
    cleanup_downloads()
    token = secrets.token_urlsafe(16)
    expires_at = datetime.utcnow() + timedelta(seconds=DOWNLOAD_TTL_SECS)
    DOWNLOADS[token] = (data, filename, media_type, expires_at)
    return token

def ensure_default_assets() -> None:
    """Crea imágenes predeterminadas si no existen"""
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
    logger.info("tiktoken disponible: recorte por tokens activo.")
except Exception:
    logger.warning("tiktoken no disponible: usando recorte aproximado por caracteres.")

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
        if "insufficient_quota" in msg or "429" in msg or "Rate limit" in msg:
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

PROMPT_ABSTRACT = """
Traduce y adapta el siguiente resumen ejecutivo del español al inglés, manteniendo la concisión, el tono profesional y la esencia del contenido. Este será el 'Abstract' del documento.

**Requisitos:**
* Traducción fiel y profesional.
* Tono: formal, claro y conciso.
* Evitar: repeticiones e información superflua.
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
# Funciones de formato DOCX
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

def format_text(p, texto, color=RGBColor(133, 78, 197)):
    tokens = re.split(r'(\*\*[^*]+\*\*|_[^_]+_|[*][^*]+[*])', texto)
    for token in tokens:
        if (token.startswith("**") and token.endswith("**")) or (token.startswith('"') and token.endswith('"')):
            run = p.add_run(token[2:-2])
            run.bold = True
        elif (token.startswith("*") and token.endswith("*")) or (token.startswith('_"') and token.endswith('"_')):
            run = p.add_run(token[1:-1])
            run.bold = True
            run.italic = True
            run.font.color.rgb = color
        elif token.startswith("_") and token.endswith("_"):
            run = p.add_run(token[1:-1])
            run.italic = True
        else:
            p.add_run(token)

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

    doc_out.add_paragraph("Tabla de contenidos", style="Heading 5")
    add_table_of_contents(doc_out.add_paragraph())

    resumen = call_gpt(api_key, PROMPT_RESUMEN, full_text[:10000], 500)
    doc_out.add_paragraph("Resumen Ejecutivo", style="Heading 1")
    doc_out.add_paragraph(resumen, style="Normal")

    abstract = call_gpt(api_key, PROMPT_ABSTRACT, resumen, 300)
    doc_out.add_paragraph("Abstract", style="Heading 1")
    doc_out.add_paragraph(abstract, style="Normal")

    hallazgos_text = full_text[:10000]
    hallazgos = call_gpt(api_key, PROMPT_HALLAZGOS, hallazgos_text, 700)
    doc_out.add_paragraph("Principales Hallazgos", style="Heading 1")
    for item in hallazgos.split("\n"):
        t = item.strip()
        if not t:
            continue
        p = doc_out.add_paragraph("", style="Normal")
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        format_text(p, t, color=REPORT_COLOR)

    for i, para in enumerate(paragraphs):
        t = para.strip()
        if not t:
            continue

        prev = paragraphs[i-1] if i > 0 else ""
        nxt  = paragraphs[i+1] if i < len(paragraphs)-1 else ""

        if t.startswith("### "):
            insert_contraportada_body(doc_out, contraportada_path)
            doc_out.add_paragraph(t[4:].strip(), style="Heading 1")
            continue
        if t.startswith("#### "):
            doc_out.add_paragraph(t[5:].strip(), style="Heading 2")
            continue
        if is_title3(prev, t, nxt):
            doc_out.add_paragraph(t.strip("*"), style="Heading 3")
            continue
        if is_title5(t):
            doc_out.add_paragraph(extract_title5_text(t), style="Heading 4")
            continue

        p = doc_out.add_paragraph("", style="Normal")
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        format_text(p, t)

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
    """Genera un informe y devuelve un link temporal para descargar"""
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
    """Versión simplificada: usa portadas/logos por defecto"""
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
    """Entrega el archivo asociado a un token válido"""
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







