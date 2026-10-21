"""
Lógica de revisión de paz y salvos (FO-AC-101) y cartas de continuidad laboral (FO-AC-103)
contra el reporte de personal FO-AC-83.

Sin dependencias de Streamlit: se puede probar desde consola
    python validador.py paz_y_salvos.pdf reporte_fo_ac_83.xlsx
"""
from __future__ import annotations

import datetime as dt
import difflib
import io
import os
import re
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass, field

import openpyxl

OK, ALERTA, ERROR, INFO = "OK", "ALERTA", "ERROR"  , "INFO"
ORDEN = {OK: 0, INFO: 0, ALERTA: 1, ERROR: 2}

# ---------------------------------------------------------------------------
# Configuración (editable desde la app)
# ---------------------------------------------------------------------------
CONFIG_DEFECTO = {
    # Asunto que debe llevar el paz y salvo FO-AC-101 según tipo de vinculación y clase de riesgo (FO-AC-83)
    "asuntos": {
        "LABORAL": "PAZ Y SALVO – PERSONAL VINCULADO POR CONTRATO LABORAL",
        "OPS_I_III": "PAZ Y SALVO – PERSONAL VINCULADO POR CONTRATO DE PRESTACIÓN DE SERVICIOS – CLASE DE RIESGO LABORAL I, II Y III",
        "OPS_IV_V": "PAZ Y SALVO – PERSONAL VINCULADO POR CONTRATO DE PRESTACIÓN DE SERVICIOS – CLASE DE RIESGO LABORAL IV y V",
    },
    "asunto_continuidad": "CERTIFICACIÓN DE CONTINUIDAD CONTRATO LABORAL",
    "tesseract_cmd": "",
    "dpi_ocr": 220,
    "una_pagina_por_documento": True,
    "solo_retiro": True,   # cruzar solo con filas RETIRO / INGRESO-RETIRO del FO-AC-83
}

MESES = {m: i for i, m in enumerate(
    ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
     "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"], 1)}

FRASES_101 = ["ARTICULO 50 DE LA LEY 789 DE 2002", "ARTICULO 23 DE LA LEY 1150 DE 2007",
              "EXAMEN MEDICO DE EGRESO", "RESOLUCIONES 2346 DE 2007", "PAZ Y SALVO"]
FRASES_103 = ["ARTICULO 50 DE LA LEY 789 DE 2002", "ARTICULO 23 DE LA LEY 1150 DE 2007",
              "NO SE EFECTUA LIQUIDACION", "OTROS PROYECTOS DE LA COMPANIA", "PAZ Y SALVO"]


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def fold(s: str) -> str:
    """Mayúsculas sin tildes ni ñ, conservando la longitud (para poder cortar el original).
    Primero junta las tildes "sueltas" (e + ´) y quita caracteres invisibles, como hacen algunos PDF y Excel."""
    s = _limpia_texto(str(s))
    out = []
    for ch in s:
        d = unicodedata.normalize("NFD", ch)
        out.append(d[0].upper() if d else ch)
    return "".join(out)


def solo_digitos(s) -> str:
    return re.sub(r"\D", "", str(s or ""))


def normaliza(s) -> str:
    return re.sub(r"\s+", " ", fold(str(s or ""))).strip()


def compacto(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", fold(str(s or "")))


def similitud(a: str, b: str) -> float:
    a, b = normaliza(a), normaliza(b)
    if not a or not b:
        return 0.0
    r1 = difflib.SequenceMatcher(None, a, b).ratio()
    r2 = difflib.SequenceMatcher(None, compacto(a), compacto(b)).ratio()
    ta, tb = sorted(a.split()), sorted(b.split())
    r3 = difflib.SequenceMatcher(None, " ".join(ta), " ".join(tb)).ratio()
    return max(r1, r2, r3)


def distancia_1(a: str, b: str) -> bool:
    """¿a y b difieren en a lo sumo 1 carácter (sustitución, inserción o borrado)?"""
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    corto, largo = sorted((a, b), key=len)
    return any(largo[:i] + largo[i + 1:] == corto for i in range(len(largo)))


STOP_CARGO = {"DE", "EN", "Y", "LA", "EL", "DEL", "A"}


def tokens_cargo(s) -> list[str]:
    t = re.sub(r"[^A-Z0-9 ]", " ", normaliza(s))
    out = []
    for w in t.split():
        if w in STOP_CARGO:
            continue
        out.append(w[:-1] + "O" if len(w) >= 6 and w.endswith("A") else w)  # Directora = Director
    return out


# ---------------------------------------------------------------------------
# Fechas y cédulas
# ---------------------------------------------------------------------------
_ABREV_MES = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AGO": 8,
              "SEP": 9, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12}


def mes_num(palabra: str):
    """Número del mes: nombre completo, abreviatura ("ENE", "SEPT") o nombre con un error de
    lectura del escáner ("FEHRERO", "SENTIEMBRE"). None si no parece un mes."""
    w = (palabra or "").upper()
    if w in MESES:
        return MESES[w]
    if 3 <= len(w) <= 5 and w[:3] in _ABREV_MES and any(m.startswith(w) for m in MESES.keys() | {"SETIEMBRE"}):
        return _ABREV_MES[w[:3]]
    if len(w) >= 5:
        c = difflib.get_close_matches(w, list(MESES), n=1, cutoff=0.75)
        if c:
            return MESES[c[0]]
    return None


def _limpiar_fecha(s: str) -> str:
    f = fold(str(s))
    f = re.sub(r"[\[\]_]", " ", f)                                   # "[2024]" (año deducido), "_03/05/2023"
    f = re.sub(r"\b2\s?\.?\s?0\s?(\d)\s?(\d)\b", r"20\1\2", f)        # "2.023", "2 024" (escáner)
    f = re.sub(r"([A-Z])(\d)", r"\1 \2", f)                             # "DE2024", "DEL2024"
    f = re.sub(r"(\d)(?=[A-Z])", r"\1 ", f)
    return re.sub(r"\s+", " ", f).strip()


def parse_fecha(s):
    """Devuelve (date|None, regular:bool). regular=False si el formato es irregular (ej. 16/072026)."""
    if s is None:
        return None, False
    if isinstance(s, dt.datetime):
        return s.date(), True
    if isinstance(s, dt.date):
        return s, True
    f = _limpiar_fecha(s)
    try:
        m = re.fullmatch(r"\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})\s*", f)
        if m:
            if int(m[2]) > 12 and int(m[1]) <= 12:   # escrita mes/día/año ("4/26/2023")
                return dt.date(int(m[3]), int(m[1]), int(m[2])), False
            return dt.date(int(m[3]), int(m[2]), int(m[1])), True
        m = re.search(r"(\d{1,2})\s+(?:D?E\s+)?([A-Z]+)\s+(?:DEL?\s+)?(\d{4})", f)
        if m and mes_num(m[2]):
            return dt.date(int(m[3]), mes_num(m[2]), int(m[1])), m[2] in MESES
        m = re.search(r"(\d{1,2})\s+DE\s+(\d{1,2})\s+DEL?\s+(\d{4})", f)
        if m:
            return dt.date(int(m[3]), int(m[2]), int(m[1])), True
        m = re.search(r"(\d{1,2})\D{0,2}(\d{2})\D{0,2}(\d{4})", f)
        if m:
            return dt.date(int(m[3]), int(m[2]), int(m[1])), False
    except ValueError:
        return None, False
    return None, False


def fecha_excel(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        return v.date() if isinstance(v, dt.datetime) else v
    if isinstance(v, (int, float)):
        try:
            return openpyxl.utils.datetime.from_excel(v).date()
        except Exception:
            return None
    d, _ = parse_fecha(v)
    return d


def fmt(d) -> str:
    return d.strftime("%d/%m/%Y") if isinstance(d, dt.date) else ""


def cedula_formato_ok(raw: str) -> bool:
    raw = (raw or "").strip()
    if not raw:
        return False
    return bool(re.fullmatch(r"\d{6,10}", raw) or re.fullmatch(r"\d{1,3}(\.\d{3}){1,3}", raw))


# ---------------------------------------------------------------------------
# Lectura del PDF (texto o OCR)
# ---------------------------------------------------------------------------
def _ocr(img, dpi_cfg: dict) -> str:
    cmd = dpi_cfg.get("tesseract_cmd") or ""
    try:
        import pytesseract
        if cmd:
            pytesseract.pytesseract.tesseract_cmd = cmd
        try:
            langs = pytesseract.get_languages(config="")
        except Exception:
            langs = ["eng"]
        lang = "spa" if "spa" in langs else "eng"
        return pytesseract.image_to_string(img, lang=lang, config="--psm 6")
    except ImportError:
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "p.png")
            img.save(ruta)
            r = subprocess.run([cmd or "tesseract", ruta, "-", "--psm", "6"],
                               capture_output=True, text=True)
            return r.stdout


class Renderizador:
    """Convierte una página del PDF en imagen. Prueba varias rutas porque algunos PDF escaneados
    se abren con pdfplumber pero PDFium los rechaza ("Data format error")."""

    def __init__(self, data: bytes):
        self.data = data
        self._reparado = None

    def _pdfium(self, datos: bytes, idx: int, dpi: int):
        import pypdfium2 as pdfium
        return pdfium.PdfDocument(datos)[idx].render(scale=dpi / 72).to_pil()

    def _reparar(self) -> bytes:
        if self._reparado is None:
            from pypdf import PdfReader, PdfWriter
            w = PdfWriter()
            for pg in PdfReader(io.BytesIO(self.data), strict=False).pages:
                w.add_page(pg)
            buf = io.BytesIO()
            w.write(buf)
            self._reparado = buf.getvalue()
        return self._reparado

    def render(self, idx: int, dpi: int = 200):
        errores = []
        for nombre, fn in (
            ("pdfplumber", lambda: __import__("pdfplumber").open(io.BytesIO(self.data)).pages[idx]
             .to_image(resolution=dpi).original),
            ("pdfium directo", lambda: self._pdfium(self.data, idx, dpi)),
            ("PDF reparado con pypdf", lambda: self._pdfium(self._reparar(), idx, dpi)),
        ):
            try:
                return fn()
            except Exception as e:  # noqa: BLE001
                errores.append(f"{nombre}: {e}")
        raise RuntimeError("No se pudo convertir la página en imagen. " + " | ".join(errores))


def leer_pdf(data: bytes, config: dict | None = None, ocr: bool = True, avisos: list | None = None) -> list[dict]:
    """Una entrada por página: {pagina, texto, fuente: 'texto'|'OCR'|'ILEGIBLE'}."""
    import pdfplumber
    config = {**CONFIG_DEFECTO, **(config or {})}
    avisos = avisos if avisos is not None else []
    rend = Renderizador(data)
    paginas = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, p in enumerate(pdf.pages):
            try:
                txt = p.extract_text() or ""
            except Exception:  # noqa: BLE001
                txt = ""
            fuente = "texto"
            if len(txt.split()) < 40 and ocr:
                try:
                    img = rend.render(i, config["dpi_ocr"]).convert("L")
                    txt = _ocr(img, config)
                    fuente = "OCR"
                except Exception as e:  # noqa: BLE001
                    fuente = "ILEGIBLE"
                    avisos.append(f"Página {i + 1}: no se pudo leer (¿escaneo dañado o Tesseract sin instalar?). {e}")
            paginas.append({"pagina": i + 1, "texto": txt, "fuente": fuente})
    return paginas


def agrupar_documentos(paginas: list[dict], una_por_pagina: bool = True) -> list[dict]:
    """Por defecto cada página es un documento (lo habitual en estos formatos).
    Con una_por_pagina=False, una página con encabezado del formato inicia un documento
    y las siguientes sin encabezado se anexan al anterior."""
    docs = []
    for p in paginas:
        ini = fold(p["texto"][:500])
        cabecera = re.search(r"FO-AC-\s?10[13]|CERTIFICACION PAZ Y SALVO|CARTA DE CONTINUIDAD", ini)
        if una_por_pagina or cabecera or not docs:
            docs.append({"paginas": [p["pagina"]], "texto": p["texto"], "fuentes": {p["fuente"]}})
        else:
            docs[-1]["paginas"].append(p["pagina"])
            docs[-1]["texto"] += "\n" + p["texto"]
            docs[-1]["fuentes"].add(p["fuente"])
    return docs


# ---------------------------------------------------------------------------
# Extracción de campos de un documento
# ---------------------------------------------------------------------------
@dataclass
class Doc:
    pagina: str = ""
    fuente: str = ""
    formato: str = ""             # FO-AC-101 / FO-AC-103 / ""
    asunto: str = ""
    nombre: str = ""
    cedula_raw: str = ""
    cedula: str = ""
    cargo: str = ""
    desde_raw: str = ""
    hasta_raw: str = ""
    hasta_contrato_raw: str = ""  # solo 103: fecha hasta la que estuvo en el contrato
    contrato: str = ""
    empresa: str = ""
    representante: str = ""
    cedula_rep: str = ""          # solo 103: cédula del representante legal (NO es la del trabajador)
    fecha_doc_raw: str = ""
    fecha_firma_raw: str = ""
    nombre_firma: str = ""
    cargo_firma: str = ""
    cedula_firma_raw: str = ""
    candidatas: int = 0       # filas del FO-AC-83 con la misma cédula entre las que se eligió
    frases_faltantes: list = field(default_factory=list)
    texto: str = ""
    es_formato: bool = True   # False = otro soporte adjunto (transferencia, etc.): no se revisa
    motivo_no_formato: str = ""  # por qué no se revisa (portada del formato, plantilla en blanco...)
    # resultado del cruce
    excel: dict | None = None
    hallazgos: list = field(default_factory=list)  # (nivel, campo, mensaje)

    def add(self, nivel, campo, msg):
        self.hallazgos.append((nivel, campo, msg))

    @property
    def estado(self):
        peor = max((ORDEN[h[0]] for h in self.hallazgos), default=0)
        return {0: OK, 1: ALERTA, 2: ERROR}[peor]


def _limpia_texto(s: str) -> str:
    """NFKC (junta tildes sueltas, ligaduras) y quita marcas combinantes / caracteres invisibles."""
    s = unicodedata.normalize("NFKC", s or "")
    return "".join(ch for ch in s if unicodedata.category(ch) not in ("Mn", "Cf"))


_MES = r"[A-Z]{3,10}"
_ANIO = r"(?:2\s?\.?\s?0\s?\d\s?\d|\d{4})"   # "2024", y lo que deja el escáner: "2.024", "2 024"
DATE = (r"(?:\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{4}"
        r"|\d{1,2}\s*(?:D?E\s+)?" + _MES + r"\s*(?:DEL?\s*)?" + _ANIO +
        r"|\d{1,2}[_\s]*DE[_\s]*\d{1,2}[_\s]*DEL?[_\s]*" + _ANIO + r"_?"
        r"|\d{1,2}\D{0,2}\d{2}\D{0,2}\d{4})")
_DIA_MES = r"\d{1,2}\s*(?:D?E\s+)?" + _MES          # fecha sin año: "22 de febrero"
_DESDE = r"DESDE\s+EL\s*(?:DIA\s*)?[_\s]*"
_HASTA = r"\s*,?\s*HASTA\s+(?:EL\s*)?(?:DIA\s*)?[_\s]*"
PERIODO = "(?:" + _DESDE + "(" + DATE + ")" + _HASTA + "(" + DATE + "))"
# Periodos donde no escribieron el año en la fecha inicial ("desde el 22 de febrero hasta el 5 de marzo
# de 2024") o en ninguna de las dos ("desde el 29 de septiembre hasta el 14 de marzo"): el año se deduce.
PERIODO_SIN_ANIO_INI = "(?:" + _DESDE + "(" + _DIA_MES + ")" + _HASTA + "(" + DATE + "))"
PERIODO_SIN_ANIO = "(?:" + _DESDE + "(" + _DIA_MES + ")" + _HASTA + "(" + _DIA_MES + r"))(?!\s*(?:D?E\s*)?\d)"


def _anio_de(raw: str):
    m = re.search(r"(\d{4})\s*$", _limpiar_fecha(raw))
    return int(m[1]) if m else None


def _con_anio(dia_mes: str, anio: int) -> str:
    """'22 de febrero' + 2024 -> '22 de febrero [2024]' (los corchetes indican que el año se dedujo)."""
    return f"{dia_mes.strip()} [{anio}]"


# Número de cédula: tolera puntos/comas en cualquier lugar ("1.012.385.270", "10123852.70") y
# espacios como separador de miles ("1 012 385 270").
# También acepta un espacio tras el último punto ("1.003.710. 049") solo si lo que sigue son exactamente 3 dígitos
# (así no se traga una fecha o un número que venga después de la cédula).
_NUM_CED = r"(\d(?:[\d\.,]|(?<=[\.,])\s(?=\d{3}(?!\d))){4,}\d|\d{1,3}(?:\s\d{3}){2,4}|\d{1,3}[\s\.,]\d{3}[\s\.,]\d{3}(?!\d)|\d{1,2}\s\d{6,8}(?!\d))"
# "cédula de ciudadanía [No. / N° / Nro. / :] <número>" (el texto ya viene sin tildes y en mayúsculas)
CED_RE = (r"(?:C[EO]DU[LT]A\s+DE\s+CIUDADAN\w*|CEDULA\s+DE\s+EXTRANJER\w*|PERMISO\s+(?:ESPECIAL\s+)?(?:DE\s+)?(?:PROTECCION\s+TEMPORAL|PERMANENCIA))\s*(?:(?:[A-Z]{1,8}\.?\s*[°º]?[\s_:]*){1,2}|[°º#:][\s_]*)?" + _NUM_CED)
CED_RE_LAXA = r"C[EO]DU[LT]A\D{0,45}?" + _NUM_CED
# En el bloque de firma: "C.C.: 123", "CC 123", "Cédula No. 123"
# Los trabajadores extranjeros firman con "PEP 4.800.980" (o PPT, pasaporte, C.E.) en vez de "C.C."
_ETIQ_DOC = r"(?:\bC\.?\s*C\b|CEDULA|\bPEP\b|\bPPT\b|PASAPORTE|\bC\.\s*E\b)"
CED_FIRMA_RE = (_ETIQ_DOC + r"\.?\s*:?\s*(?:N[O°º]?\.?\s*)?[°º#:]?\s*" + _NUM_CED)


def _buscar_cedula(t: str, f: str, desde: int = 0, hasta: int | None = None, ultima: bool = False) -> str:
    """Número de cédula (tal como está escrito) que sigue a 'cédula de ciudadanía' dentro de f[desde:hasta].
    t es el texto original y f su versión en mayúsculas sin tildes (misma longitud)."""
    fin = len(f) if hasta is None else hasta
    ms = list(re.finditer(CED_RE, f[desde:fin])) or list(re.finditer(CED_RE_LAXA, f[desde:fin]))
    if not ms:
        return ""
    m = ms[-1] if ultima else ms[0]
    return t[desde + m.start(1): desde + m.end(1)].strip(" ,.;:")


MARCAS_PORTADA = ("CONTROLDEVERSIONES", "PARTICIPARONENLAELABORACION", "DESCRIPCIONMODIFICACION",
                  "VALIDADOPOR", "REVISADOPOR", "APROBADOPOR", "SISTEMAINFORMACIONDOCUMENTADA",
                  "AUTENTICIDADPUEDESER", "ELDOCUMENTOORIGINALHASIDO", "ALCANCEDEPARTICIPACION",
                  "VALIDADOEL", "REVISADOEL", "APROBADOEL")


def _marcar_portada_o_plantilla(d: "Doc", f: str):
    """Páginas que NO son un certificado aunque digan FO-AC-101 / FO-AC-103:
    - la portada / instructivo del formato ("Control de Versiones", "Participaron en
      la elaboración", "Aprobado por", código QR del SID);
    - la plantilla sin diligenciar ("Escriba el nombre del trabajador", "Escriba el
      número del día"...);
    - una página que solo trae el encabezado del formato, sin texto del certificado.
    Nunca se descarta una página que traiga el texto de un certificado diligenciado."""
    fs = re.sub(r"[^A-Z]", "", f)
    cuerpo = ("POR MEDIO DE LA PRESENTE" in f or "IDENTIFICAD" in f
              or "EN CALIDAD DE REPRESENTANTE" in f or "CERTIFICO" in f)
    marcas_portada = sum(k in fs for k in MARCAS_PORTADA)
    marcas_plantilla = len(re.findall(r"\bESCRIBA\s+(?:EL|LA|LOS|LAS|AQUI|SU)\b", f))
    nombre_real = d.nombre and "ESCRIBA" not in fold(d.nombre)
    motivo = ""
    if marcas_portada >= 2 and not ("POR MEDIO DE LA PRESENTE" in f or "IDENTIFICAD" in f):
        motivo = "portada / instructivo del formato (Control de versiones): no tiene datos de un trabajador"
    elif marcas_plantilla >= 2 and not nombre_real:
        motivo = "plantilla del formato sin diligenciar (textos 'Escriba el...')"
    elif "CERTIFICO" not in f and "IDENTIFICAD" not in f and not d.cedula_raw and not nombre_real:
        if "LIQUIDACION" in f and ("VOLANTE DE PAGO" in f or "DEVENGAD" in f):
            motivo = "soporte: volante de pago / liquidación final (no es un FO-AC-101 ni FO-AC-103)"
        elif "NO ES POSIBLE ALLEGAR" in f or "NO SE PUEDE ALLEGAR" in f:
            motivo = "carta que informa que NO se entrega el paz y salvo (revisar el motivo en el documento)"
        else:
            motivo = "soporte o carta adjunta (no es un FO-AC-101 ni FO-AC-103)"
    elif d.formato and not cuerpo and len(f.split()) < 60:
        motivo = "página con solo el encabezado del formato, sin datos de un trabajador"
    if motivo:
        d.es_formato = False
        d.motivo_no_formato = motivo


def _periodo_sin_anio(d: "Doc", t: str, f: str):
    """Periodo escrito sin año en la fecha inicial (o en ambas). El año se deduce de la fecha final
    o, si tampoco lo tiene, de la fecha de desvinculación (103) o de la fecha del documento.
    El año deducido queda entre corchetes para avisarlo en el reporte."""
    ultimo = d.formato == "FO-AC-103"
    ms = list(re.finditer(PERIODO_SIN_ANIO_INI, f))
    if ms:
        m = ms[-1] if ultimo else ms[0]
        ini, fin = t[m.start(1):m.end(1)].strip(), t[m.start(2):m.end(2)].strip()
        hasta, _ = parse_fecha(fin)
        mi = re.search(r"([A-Z]+)\s*$", fold(ini))
        if hasta and mi and mes_num(mi[1]):
            anio = hasta.year - (1 if mes_num(mi[1]) > hasta.month else 0)
            d.desde_raw, d.hasta_raw = _con_anio(ini, anio), fin
        return
    ms = list(re.finditer(PERIODO_SIN_ANIO, f))
    if not ms:
        return
    m = ms[-1] if ultimo else ms[0]
    ini, fin = t[m.start(1):m.end(1)].strip(), t[m.start(2):m.end(2)].strip()
    ref, _ = parse_fecha(d.hasta_contrato_raw)
    if ref is None:
        ref, _ = parse_fecha(d.fecha_doc_raw)
    mi, mf = re.search(r"([A-Z]+)\s*$", fold(ini)), re.search(r"([A-Z]+)\s*$", fold(fin))
    if not (ref and mi and mf and mes_num(mi[1]) and mes_num(mf[1])):
        return
    a_fin = ref.year - (1 if mes_num(mf[1]) > ref.month else 0)
    a_ini = a_fin - (1 if mes_num(mi[1]) > mes_num(mf[1]) else 0)
    d.desde_raw, d.hasta_raw = _con_anio(ini, a_ini), _con_anio(fin, a_fin)


def extraer_campos(texto: str, paginas, fuentes) -> Doc:
    t = re.sub(r"\s+", " ", _limpia_texto(texto)).strip()
    f = fold(t)
    d = Doc(pagina=",".join(map(str, paginas)), fuente="/".join(sorted(fuentes)), texto=t)

    def g(pat, grp=1, flags=0):
        m = re.search(pat, f, flags)
        return t[m.start(grp):m.end(grp)].strip(" ,.;:") if m else ""

    m = re.search(r"FO-AC-\s?(101|103)", f)
    if m:
        d.formato = f"FO-AC-{m[1]}"
    elif "CONTINUIDAD" in f[:400]:
        d.formato = "FO-AC-103"
    elif "PAZ Y SALVO" in f[:400]:
        d.formato = "FO-AC-101"

    # ¿Es realmente un paz y salvo / carta de continuidad? Si no, es un soporte adjunto y no se revisa.
    d.es_formato = bool(d.formato) or (
        ("PAZ Y SALVO" in f or "CONTINUIDAD LABORAL" in f)
        and ("IDU" in f or "CONTRATO" in f or "CERTIFICO" in f))

    mp = list(re.finditer(PERIODO, f))
    if d.formato == "FO-AC-103":
        d.nombre = g(r"CERTIFICO(?:\s+QUE)?\s+(.+?)\s*,?\s+IDENTIFICAD\w*\s+CON")
        d.cargo = g(r"SE DESEMPE\w{0,8}\s+COMO\s+(.+?)\s*,?\s+EN EL CONTRATO")
        if mp:  # en la carta de continuidad el periodo de paz y salvo es el último "desde ... hasta ..."
            d.desde_raw = t[mp[-1].start(1):mp[-1].end(1)].strip()
            d.hasta_raw = t[mp[-1].start(2):mp[-1].end(2)].strip()
        d.hasta_contrato_raw = g(r"HASTA\s+EL\s*(?:DIA\s+)?(" + DATE + r")\s*,?\s*SEGUIRA")
        d.empresa = g(r"REPRESENTANTE LEGAL(?: Y/O REVISOR FISCAL)? DE\s+(.+?)\s+PARA EL CONTRATO")
        d.representante = g(r"\bYO\s+(.+?)\s+EN CALIDAD DE REPRESENTANTE")
    else:  # 101 (o no identificado)
        d.nombre = g(r"POR MEDIO DE LA PRESENTE\s+(?:YO,?\s+)?(.+?)\s*,?\s+IDENTIFICAD\w*\s+CON")
        d.cargo = g(r"EN CALIDAD\.?\s+DE[_\s]+(.+?)\s*DESDE\s+EL")
        if not mp:  # sin "desde el día": "en calidad de AYUDANTE ... 09 de enero 2024 hasta el 19 de febrero 2024"
            m2 = re.search(r"EN CALIDAD\.?\s+DE[_\s]+((?:(?!DESDE)[^\d]){3,80}?)\s+(" + DATE + r")"
                           + _HASTA + "(" + DATE + ")", f)
            if m2:
                d.cargo = d.cargo or t[m2.start(1):m2.end(1)].strip()
                d.desde_raw = t[m2.start(2):m2.end(2)].strip()
                d.hasta_raw = t[m2.start(3):m2.end(3)].strip()
        if mp:
            d.desde_raw = t[mp[0].start(1):mp[0].end(1)].strip()
            d.hasta_raw = t[mp[0].start(2):mp[0].end(2)].strip()
        d.empresa = g(r"PERMITO CERTIFICAR QUE\s+(?:EL\s+|LA\s+)?(.+?)\s*,?\s+EST[AE]\s+A\s+PAZ")

    d.nombre = re.sub(r"^yo,?\s*(?=[A-ZÁÉÍÓÚÑ])", "", d.nombre)  # "yoLUIS ..." (palabras pegadas)
    d.asunto = g(r"ASUNTO:\s*(.+?)\s+RESPETADOS")
    if d.formato == "FO-AC-103":
        # La primera cédula del texto es la del representante legal; la del trabajador va después de "CERTIFICO QUE".
        mq = re.search(r"\bCERTIFICO\b", f)
        if mq:
            d.cedula_rep = _buscar_cedula(t, f, 0, mq.start())
            d.cedula_raw = _buscar_cedula(t, f, mq.end())
        else:
            d.cedula_raw = _buscar_cedula(t, f, ultima=True)
    else:
        d.cedula_raw = _buscar_cedula(t, f)
    d.cedula = solo_digitos(d.cedula_raw)
    mc = re.search(r"CONTRATO IDU\s*(?:NO|N°|N)\.?\s*(\d{2,4})\s*-\s*(\d{4})", f)
    d.contrato = f"{mc[1]}-{mc[2]}" if mc else ""

    d.fecha_doc_raw = g(r"BOGOTA[,\.]?\s*D\.?\s*[C€]\.?\s*[_\s]*(" + DATE + r")")
    d.fecha_firma_raw = g(r"SE FIRMA\s+(?:EN\s+[A-Z\. ]+?\s+)?A LOS,?\s*"
                          r"(\d{1,2}\s+DIA[S]?(?:\(S\))?\s+DEL MES\s+(?:DE\s+)?[A-Z]+\s+(?:DEL\s+\S+\s+|DE\s+)\d{4})")

    if not d.desde_raw and not d.hasta_raw:
        _periodo_sin_anio(d, t, f)

    # Bloque de firmas (el último "Nombre:" es el trabajador)
    nombres = [t[m.start(1):m.end(1)].strip() for m in
               re.finditer(r"NOMBRE:?\s*([A-Z ]+?)(?=\s+(?:NOMBRE|REPRESENTANTE|TRABAJADOR|CONTRATISTA|CARGO|PEP|PPT|C\.\s*C))", f)]
    d.nombre_firma = nombres[-1] if nombres else ""
    d.cargo_firma = g(r"CARGO:\s*(.+?)\s+C\.\s*C")
    ceds = [t[m.start(1):m.end(1)].strip(" ,.;:") for m in re.finditer(CED_FIRMA_RE, f)]
    d.cedula_firma_raw = ceds[-1] if ceds else ""

    # Frases de la plantilla (solo contrato laboral y carta de continuidad; los formatos OPS cambian el texto)
    clave = clave_asunto(d.asunto)
    base = FRASES_103 if d.formato == "FO-AC-103" else FRASES_101 if clave == "LABORAL" else []
    fc = compacto(f)
    d.frases_faltantes = [x for x in base if compacto(x) not in fc]
    _marcar_portada_o_plantilla(d, f)
    return d


# ---------------------------------------------------------------------------
# Bloque de firma por posición (carta de continuidad FO-AC-103)
# ---------------------------------------------------------------------------
# En la carta hay dos bloques "Nombre / Cargo / C.C.": el del representante legal y el del trabajador
# (abajo a la derecha). El orden del texto extraído no es confiable (según la altura de cada bloque
# puede quedar primero o último), así que se localiza por coordenadas: el bloque más a la derecha y más abajo.
def palabras_pdf(data: bytes, idx: int):
    """Palabras de la página con sus coordenadas (PDF con texto). Devuelve (palabras, ancho, alto)."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        pg = pdf.pages[idx]
        ws = pg.extract_words(x_tolerance=2, y_tolerance=3)
        return ([{"t": w["text"], "x0": w["x0"], "top": w["top"], "bottom": w["bottom"]} for w in ws],
                float(pg.width), float(pg.height))


def palabras_ocr(rend: "Renderizador", idx: int, cfg: dict, desde_frac: float = 0.40):
    """Palabras con coordenadas por OCR, solo de la parte baja de la página (donde van las firmas)."""
    import pytesseract
    img = rend.render(idx, cfg["dpi_ocr"]).convert("L")
    ancho, alto = img.size
    y0 = int(alto * desde_frac)
    if cfg.get("tesseract_cmd"):
        pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_cmd"]
    try:
        langs = pytesseract.get_languages(config="")
    except Exception:  # noqa: BLE001
        langs = ["eng"]
    dd = pytesseract.image_to_data(img.crop((0, y0, ancho, alto)), lang="spa" if "spa" in langs else "eng",
                                   config="--psm 11", output_type=pytesseract.Output.DICT)
    ws = []
    for txt, l, tp, h, conf in zip(dd["text"], dd["left"], dd["top"], dd["height"], dd["conf"]):
        if str(txt).strip() and float(conf) >= 0:
            ws.append({"t": str(txt).strip(), "x0": float(l), "top": float(tp + y0), "bottom": float(tp + y0 + h)})
    return ws, float(ancho), float(alto)


def ocr_region(rend: "Renderizador", idx: int, cfg: dict, caja: tuple) -> str:
    """OCR de un recuadro de la página (x0, y0, x1, y1 en píxeles). Lee líneas completas: más fiable que el modo
    de texto disperso cuando hay una firma manuscrita encima del nombre."""
    import pytesseract
    img = rend.render(idx, cfg["dpi_ocr"]).convert("L")
    x0, y0, x1, y1 = (int(max(c, 0)) for c in caja)
    if cfg.get("tesseract_cmd"):
        pytesseract.pytesseract.tesseract_cmd = cfg["tesseract_cmd"]
    try:
        langs = pytesseract.get_languages(config="")
    except Exception:  # noqa: BLE001
        langs = ["eng"]
    return pytesseract.image_to_string(img.crop((x0, y0, min(x1, img.size[0]), min(y1, img.size[1]))),
                                       lang="spa" if "spa" in langs else "eng", config="--psm 6")


def _texto_en_lineas(ws: list[dict]) -> str:
    """Une las palabras en orden de lectura (por líneas, de arriba abajo y de izquierda a derecha)."""
    ws = sorted(ws, key=lambda w: (w["top"], w["x0"]))
    lineas, cur, ref = [], [], 0.0
    for w in ws:
        alto = max(w["bottom"] - w["top"], 1.0)
        if cur and abs(w["top"] - ref) > alto * 0.6:
            lineas.append(cur)
            cur = []
        if not cur:
            ref = w["top"]
        cur.append(w)
    if cur:
        lineas.append(cur)
    return " ".join(" ".join(x["t"] for x in sorted(l, key=lambda w: w["x0"])) for l in lineas)


def parsear_bloque(texto: str) -> dict:
    """De 'Nombre: X Cargo: Y C.C.: Z' saca nombre, cargo y cédula (tal como están escritos)."""
    bt = re.sub(r"\s+", " ", _limpia_texto(texto)).strip()
    bt = re.sub(r"\bN\s?o\s?m\s?b\s?r\s?e(?=\s?:)", "Nombre", bt, flags=re.I)  # "N o m bre:" (letras partidas en el PDF)
    bf = fold(bt)
    out = {"nombre": "", "cargo": "", "cedula": ""}
    m = re.search(r"NOMBRES?:?\s*([A-Z ]+?)(?=\s+(?:CARGO|C\.?\s*C\b|CEDULA|PEP|PPT|PASAPORTE|REPRESENTANTE|TRABAJADOR|CONTRATISTA|FIRMA|NOMBRE)\b|\s*\d|$)", bf)
    if m:
        out["nombre"] = bt[m.start(1):m.end(1)].strip(" ,.;:")
    m = re.search(r"CARGO:?\s*(.+?)(?=\s+(?:C\.?\s*C\b|CEDULA|NOMBRE)\b|$)", bf)
    if m:
        out["cargo"] = bt[m.start(1):m.end(1)].strip(" ,.;:")
    ms = list(re.finditer(CED_FIRMA_RE, bf))
    if ms:
        out["cedula"] = bt[ms[-1].start(1):ms[-1].end(1)].strip(" ,.;:")
    return out


def _anclas_nombre(palabras: list[dict], ancho: float) -> list[dict]:
    """Palabras que inician una etiqueta 'Nombre:'. Algunos PDF la parten en piezas ('N', 'o', 'm', 'bre:'),
    así que se reúnen las palabras contiguas de la misma línea antes de decidir."""
    ws = sorted(palabras, key=lambda w: (w["top"], w["x0"]))
    anclas = []
    for i, w in enumerate(ws):
        junto = ""
        ult = w
        for w2 in ws[i:i + 6]:
            if abs(w2["top"] - w["top"]) > max(w["bottom"] - w["top"], 1.0) * 0.6 or w2["x0"] - ult["x0"] > ancho * 0.08:
                break
            junto += re.sub(r"[^A-Z]", "", fold(w2["t"]))
            ult = w2
            if junto.startswith("NOMBRE"):
                anclas.append(w)
                break
            if not "NOMBRE".startswith(junto):
                break
    return anclas


def bloques_firma(palabras: list[dict], ancho: float, alto: float) -> list[dict]:
    """Bloques 'Nombre/Cargo/C.C.' de la parte baja de la página, del más a la derecha y más abajo al más a la izquierda."""
    anclas = _anclas_nombre([w for w in palabras if w["top"] > alto * 0.40], ancho)
    bloques = []
    for a in anclas:
        derecha = [b["x0"] for b in anclas if b is not a and b["x0"] > a["x0"] + ancho * 0.05
                   and abs(b["top"] - a["top"]) < alto * 0.12]
        debajo = [b["top"] for b in anclas if b is not a and abs(b["x0"] - a["x0"]) < ancho * 0.05
                  and b["top"] > a["top"] + 3]
        x_fin = (min(derecha) - 2) if derecha else float("inf")
        y_fin = min(a["top"] + alto * 0.16, (min(debajo) - 1) if debajo else float("inf"))
        sel = [w for w in palabras if a["x0"] - ancho * 0.02 <= w["x0"] < x_fin and a["top"] - 3 <= w["top"] <= y_fin]
        txt = _texto_en_lineas(sel)
        if not re.search(_ETIQ_DOC + r"|CARGO|TRABAJADOR|REPRESENTANTE", fold(txt)):
            continue  # la palabra "nombre" del cuerpo del texto, no una etiqueta de firma
        b = parsear_bloque(txt)
        if b["nombre"] or b["cedula"]:
            ft = fold(txt)
            b.update(x0=a["x0"], top=a["top"], rep="REPRESENTANTE" in ft, trab="TRABAJADOR" in ft or "CONTRATISTA" in ft,
                     caja=(a["x0"] - ancho * 0.015, a["top"] - alto * 0.004, min(x_fin, ancho),
                           min(y_fin, a["top"] + alto * 0.09)))
            bloques.append(b)
    if not bloques:
        return []
    xmax = max(b["x0"] for b in bloques)
    # primero el que dice "Trabajador o contratista"; luego el que no dice "Representante"; luego el más abajo a la derecha
    return sorted(bloques, key=lambda b: (not b["trab"], b["rep"], b["x0"] < xmax - ancho * 0.12, -b["top"], -b["x0"]))


def _es_representante(nombre: str, rep_nombre: str) -> bool:
    """¿Este nombre es el del representante legal? (igual, casi igual o el mismo con algún apellido de menos)"""
    if similitud(nombre, rep_nombre) >= 0.9:
        return True
    a, b = set(normaliza(nombre).split()), set(normaliza(rep_nombre).split())
    return min(len(a), len(b)) >= 2 and (a <= b or b <= a)


def aplicar_bloque_firma(d: Doc, data: bytes, idx: int, cfg: dict, rend: "Renderizador",
                         usar_ocr: bool, avisos: list, etiqueta: str = ""):
    """Carta FO-AC-103: toma nombre, cargo y cédula del trabajador del bloque de firma de abajo a la derecha."""
    por_ocr = False
    try:
        ws, ancho, alto = palabras_pdf(data, idx)
        if len(ws) < 40 and usar_ocr:
            ws, ancho, alto = palabras_ocr(rend, idx, cfg)
            por_ocr = True
        cands = bloques_firma(ws, ancho, alto)
    except Exception as e:  # noqa: BLE001
        avisos.append(f"{etiqueta} · pág. {d.pagina}: no se pudo ubicar el bloque de firma por posición ({e}).")
        return
    rep_nombre = normaliza(re.split(r"IDENTIFICAD", fold(d.representante))[0]) if d.representante else ""
    elegido = None
    for c in cands:  # nunca el bloque del representante legal
        if c.get("rep") and not c.get("trab"):
            continue
        if rep_nombre and c["nombre"] and _es_representante(c["nombre"], rep_nombre):
            continue
        if d.cedula_rep and c["cedula"] and solo_digitos(c["cedula"]) == solo_digitos(d.cedula_rep):
            continue
        elegido = c
        break
    if not elegido:
        if rep_nombre and d.nombre_firma and _es_representante(d.nombre_firma, rep_nombre):
            d.nombre_firma = ""        # lo que se leyó como "firma" es el representante legal: se descarta
            d.cedula_firma_raw = ""    # y su cédula también
            d.cargo_firma = ""
        return
    if por_ocr and elegido.get("caja"):
        try:  # segunda lectura, solo del recuadro del bloque, con OCR de líneas
            b2 = parsear_bloque(ocr_region(rend, idx, cfg, elegido["caja"]))
            if len(b2["nombre"].split()) >= len(elegido["nombre"].split()):
                elegido["nombre"] = b2["nombre"] or elegido["nombre"]
            elegido["cedula"] = b2["cedula"] or elegido["cedula"]
            elegido["cargo"] = b2["cargo"] or elegido["cargo"]
        except Exception:  # noqa: BLE001
            pass
    if elegido["nombre"]:
        d.nombre_firma = elegido["nombre"]
    if elegido["cedula"]:
        d.cedula_firma_raw = elegido["cedula"]
    if elegido["cargo"]:
        d.cargo_firma = elegido["cargo"]
    if d.nombre_firma and (not d.nombre or (rep_nombre and _es_representante(d.nombre, rep_nombre))):
        d.nombre = d.nombre_firma
    if not d.cedula_raw and d.cedula_firma_raw:
        d.cedula_raw = d.cedula_firma_raw
        d.cedula = solo_digitos(d.cedula_raw)


CAMPOS_CLAVE = ("nombre", "cedula", "cargo", "desde_raw", "hasta_raw", "fecha_doc_raw", "empresa")


def sospechoso(valor: str, largo: int) -> bool:
    """Palabras pegadas por el PDF (ej. 'yoLUIS AURELIOPAEZ'): alguna palabra demasiado larga."""
    return any(len(w) > largo for w in (valor or "").split())


def campos_faltantes(d: Doc) -> list[str]:
    falta = [c for c in CAMPOS_CLAVE if not getattr(d, c)]
    if d.nombre and d.nombre_firma and compacto(d.nombre) != compacto(d.nombre_firma):
        ta, tb = set(normaliza(d.nombre).split()), set(normaliza(d.nombre_firma).split())
        if not (ta <= tb or tb <= ta):
            falta.append("nombre")  # pueden ser palabras pegadas: se relee la página
    for c, largo in (("nombre", 15), ("nombre_firma", 15), ("cargo", 20), ("empresa", 20)):
        if sospechoso(getattr(d, c), largo):
            falta.append(c)
    if d.formato != "FO-AC-103" and (not d.asunto or clave_asunto(d.asunto) == "OPS"):
        falta.append("asunto")  # vacío o cortado (sin la clase de riesgo)
    if d.formato == "FO-AC-103" and not d.hasta_contrato_raw:
        falta.append("hasta_contrato_raw")
    if d.formato == "FO-AC-103" and not d.asunto:
        falta.append("asunto")
    if d.frases_faltantes:
        falta.append("plantilla")  # la capa de texto de un escaneo a veces pierde frases: se relee por OCR
    for c in ("desde_raw", "hasta_raw"):
        if getattr(d, c) and parse_fecha(getattr(d, c))[0] is None:
            falta.append(c)
    return falta


def fusionar(d: Doc, d2: Doc):
    """Completa los campos vacíos (o ilegibles) de d con los leídos en d2 (otra lectura de la misma página)."""
    for c in ("nombre", "cedula_raw", "cargo", "desde_raw", "hasta_raw", "hasta_contrato_raw", "empresa",
              "representante", "fecha_doc_raw", "fecha_firma_raw", "nombre_firma", "cargo_firma",
              "cedula_firma_raw", "formato"):
        a, b = getattr(d, c), getattr(d2, c)
        if not a and b:
            setattr(d, c, b)
        elif a and b and c in ("nombre", "nombre_firma", "cargo", "empresa") and compacto(a) == compacto(b) \
                and len(b.split()) > len(a.split()):
            setattr(d, c, b)
        elif a and b and c in ("nombre", "nombre_firma", "cargo", "empresa") and sospechoso(a, 15 if "nombre" in c else 20) \
                and not sospechoso(b, 15 if "nombre" in c else 20):
            setattr(d, c, b)
        elif a and b and c in ("desde_raw", "hasta_raw", "hasta_contrato_raw", "fecha_doc_raw"):
            if parse_fecha(a)[0] is None and parse_fecha(b)[0] is not None:
                setattr(d, c, b)
    d.cedula = solo_digitos(d.cedula_raw)
    if (not d.asunto) or (clave_asunto(d.asunto) in ("", "OPS") and clave_asunto(d2.asunto) not in ("", "OPS")):
        d.asunto = d2.asunto or d.asunto
    d.texto += " " + d2.texto
    d.frases_faltantes = [p for p in d.frases_faltantes if p in d2.frases_faltantes]


def texto_pagina(data: bytes, idx: int, modo: str, cfg: dict, rend: "Renderizador") -> str:
    """Lectura alterna de una página: 'xtol' (texto con menor tolerancia entre letras) u 'ocr'."""
    if modo == "xtol":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return pdf.pages[idx].extract_text(x_tolerance=1.5) or ""
    return _ocr(rend.render(idx, cfg["dpi_ocr"]).convert("L"), cfg)


# ---------------------------------------------------------------------------
# Lectura del FO-AC-83
# ---------------------------------------------------------------------------
def _ced_celda(v) -> str:
    """Dígitos de la cédula de una celda; evita que 1012385270.0 (número guardado como decimal) se lea con un 0 de más."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"^(\d+)\.0+$", r"\1", str(v if v is not None else "").strip())
    return solo_digitos(s)


def cargar_fo83(origen) -> dict:
    """Devuelve {'filas': [...], 'periodo': (desde, hasta), 'contrato': str, 'contratista': str}."""
    wb = openpyxl.load_workbook(origen, data_only=True)
    ws = None
    cols = {}
    fila_enc = None
    for hoja in wb.worksheets:
        for r in range(1, 60):
            for c in range(1, 70):
                v = hoja.cell(r, c).value
                if isinstance(v, str) and normaliza(v) == "CEDULA":
                    ws, fila_enc = hoja, r
                    break
            if ws:
                break
        if ws:
            break
    if ws is None:
        raise ValueError("No encontré la tabla de personal (columna CEDULA) en el Excel.")

    etiquetas = {"TIPO DE CONTRATO": "tipo", "NOMBRE DEL TRABAJADOR": "nombre", "CEDULA": "cedula",
                 "CARGO": "cargo", "FECHA INGRESO": "ingreso", "FECHA RETIRO": "retiro", "CLASE": "clase"}
    for c in range(1, 70):
        for r in (fila_enc, fila_enc - 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            n = normaliza(v)
            if n in etiquetas and etiquetas[n] not in cols:
                cols[etiquetas[n]] = c
            if n.startswith("EMPRESA") and "empresa" not in cols:
                cols["empresa"] = c
            if n == "NOMBRE" and "retiro" in cols and c > cols["retiro"] and "arl" not in cols:
                cols["arl"] = c
    faltan = [k for k in ("nombre", "cedula", "cargo", "ingreso", "retiro", "tipo") if k not in cols]
    if faltan:
        raise ValueError(f"Faltan columnas en el FO-AC-83: {faltan}")

    filas, seccion, grupo = [], "", ""
    for r in range(fila_enc + 1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.strip() and not a.strip().isdigit():
            n = normaliza(a)
            if "ROTATIVO" in n:
                grupo, seccion = "ROTATIVO", ""
            else:
                seccion = n
            continue
        ced = _ced_celda(ws.cell(r, cols["cedula"]).value)
        nombre = ws.cell(r, cols["nombre"]).value
        if len(ced) < 5 or not nombre:
            continue
        clase = str(ws.cell(r, cols["clase"]).value or "").strip().upper() if "clase" in cols else ""
        clase = {"1": "I", "2": "II", "3": "III", "4": "IV", "5": "V"}.get(clase, clase)
        filas.append({
            "fila_excel": r,
            "seccion": (f"{grupo} / {seccion}" if grupo else seccion),
            "empresa": str(ws.cell(r, cols["empresa"]).value or "").strip() if "empresa" in cols else "",
            "tipo": normaliza(ws.cell(r, cols["tipo"]).value),
            "nombre": str(nombre).strip(),
            "cedula": ced,
            "cargo": str(ws.cell(r, cols["cargo"]).value or "").strip(),
            "ingreso": fecha_excel(ws.cell(r, cols["ingreso"]).value),
            "retiro": fecha_excel(ws.cell(r, cols["retiro"]).value),
            "arl": str(ws.cell(r, cols["arl"]).value or "").strip() if "arl" in cols else "",
            "clase": clase,
        })

    # Encabezado: contrato y periodo del informe
    contrato, contratista, desde, hasta = "", "", None, None
    for r in range(1, fila_enc):
        vals = [(c, ws.cell(r, c).value) for c in range(1, 70) if ws.cell(r, c).value not in (None, "")]
        for i, (c, v) in enumerate(vals):
            if not isinstance(v, str):
                continue
            n = normaliza(v)
            sig = vals[i + 1][1] if i + 1 < len(vals) else None
            if n.startswith("CONTRATO NO") and not contrato:
                contrato = str(sig or "")
            if n.startswith("CONTRATISTA /") and not contratista:
                contratista = str(sig or "")
            if n == "DESDE" and desde is None:
                desde = fecha_excel(sig)
            if n == "HASTA" and hasta is None:
                hasta = fecha_excel(sig)
    return {"filas": filas, "periodo": (desde, hasta), "contrato": contrato, "contratista": contratista}


# ---------------------------------------------------------------------------
# Validación
# ---------------------------------------------------------------------------
MIN_DIGITOS_APROX = 5  # dígitos de la cédula que deben coincidir para un cruce aproximado
MIN_NOMBRES_APROX = 1  # nombres/apellidos que deben coincidir cuando la cédula difiere en 2 dígitos o menos
# Si la cédula difiere en más de 2 dígitos (pero aún coinciden MIN_DIGITOS_APROX) se piden más nombres/apellidos,
# porque las cédulas de una misma zona comparten los primeros dígitos (80.267.270 y 80.267.999) y un solo nombre
# común (CARLOS) no basta. Con MIN_NOMBRES_DEBIL = 1 se aplica exactamente "5 dígitos + un nombre o apellido".
MIN_NOMBRES_DEBIL = 2
_NOMBRE_STOP = {"DE", "DEL", "LA", "LAS", "LOS", "Y", "E"}


def _fold_token(t: str) -> str:
    """Clave fonética simple: Y/J/I, B/V, S/Z, K/C, H muda y letras dobles se consideran iguales (YENNY = JENNY)."""
    t = re.sub(r"[^A-Z]", "", fold(t))
    t = t.replace("H", "").replace("V", "B").replace("Z", "S").replace("K", "C").replace("J", "I").replace("Y", "I")
    return re.sub(r"(.)\1+", r"\1", t)


def _palabras_nombre(s: str) -> list[str]:
    return [_fold_token(w) for w in normaliza(s).split() if len(w) >= 3 and w not in _NOMBRE_STOP]


def coincidencias_nombre(a: str, b: str) -> int:
    """Cuántos nombres/apellidos de a aparecen, iguales o casi iguales (1 letra), en b."""
    tb = _palabras_nombre(b)
    hits = 0
    for x in _palabras_nombre(a):
        if x and any(x == y or (len(x) >= 4 and len(y) >= 4 and difflib.SequenceMatcher(None, x, y).ratio() >= 0.85)
                     for y in tb):
            hits += 1
    return hits


def digitos_comunes(a: str, b: str) -> int:
    """Dígitos que coinciden entre dos cédulas: tolera un dígito cambiado (misma posición), uno de más o de menos
    en cualquier lugar (prefijo + sufijo común) o corrimiento por el final."""
    if len(a) < 5 or len(b) < 5 or abs(len(a) - len(b)) > 2:
        return 0
    izq = sum(x == y for x, y in zip(a, b))
    der = sum(x == y for x, y in zip(reversed(a), reversed(b)))
    pre = 0
    while pre < min(len(a), len(b)) and a[pre] == b[pre]:
        pre += 1
    ra, rb = a[pre:], b[pre:]
    suf = 0
    while suf < min(len(ra), len(rb)) and ra[-1 - suf] == rb[-1 - suf]:
        suf += 1
    return max(izq, der, pre + suf)


def buscar_aproximada(d: Doc, ceds_doc: list[str], filas: list[dict]) -> list[dict]:
    """Personas del FO-AC-83 con >= MIN_DIGITOS_APROX dígitos de cédula en común Y al menos un nombre/apellido
    igual o parecido. Una entrada por cédula, de la mejor a la peor."""
    nombres_doc = [n for n in (d.nombre_firma, d.nombre) if n]
    mejores: dict[str, dict] = {}
    for x in filas:
        ds, ced_doc = max(((digitos_comunes(c, x["cedula"]), c) for c in ceds_doc), default=(0, ""))
        if ds < MIN_DIGITOS_APROX:
            continue
        hits = max((coincidencias_nombre(n, x["nombre"]) for n in nombres_doc), default=0)
        largo = min(len(ced_doc), len(x["cedula"]))
        if hits < (MIN_NOMBRES_APROX if ds >= largo - 2 else MIN_NOMBRES_DEBIL):
            continue
        cand = {"cedula": x["cedula"], "nombre": x["nombre"], "ds": ds, "hits": hits, "ced_doc": ced_doc}
        if x["cedula"] not in mejores or (ds, hits) > (mejores[x["cedula"]]["ds"], mejores[x["cedula"]]["hits"]):
            mejores[x["cedula"]] = cand
    return sorted(mejores.values(), key=lambda c: (-c["ds"], -c["hits"]))


def _es_mda(raw: str) -> bool:
    """Fecha escrita mes/día/año (el segundo número es mayor que 12): '3/20/2024'."""
    m = re.fullmatch(r"[\s_]*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*\d{4}[\s_]*", raw or "")
    return bool(m) and int(m[2]) > 12 and int(m[1]) <= 12


def _es_cedula_representante(d: Doc, ced: str) -> bool:
    """¿Esta cédula es la del representante legal? (la que va con 'Representante legal' en las firmas)."""
    c = solo_digitos(ced)
    if not c:
        return False
    reps = [solo_digitos(x) for x in (d.cedula_rep,) if x]
    f = fold(d.texto or "")
    for m in re.finditer(r"REPRESENTANTE\s+LEGAL", f):
        mc = re.search(CED_FIRMA_RE, f[m.end():m.end() + 120])
        if mc:
            reps.append(solo_digitos(mc[1]))
    return any(r and (r == c or (len(r) >= 7 and difflib.SequenceMatcher(None, r, c).ratio() >= 0.8)) for r in reps)


def _firma_legible(nombre: str, firma: str) -> bool:
    """El nombre leído en la firma tiene al menos una palabra parecida al nombre del texto
    (si no, es casi seguro letra a mano que el OCR no entendió: 'Jerye a', 'D Nombr')."""
    pf = [w for w in normaliza(firma).split() if len(w) >= 3 and not w.startswith(("NOMBR", "FIRMA"))]
    if len(pf) < 2:
        return False   # una sola palabra ('Luis', 'obio'): lectura incompleta
    pn = normaliza(nombre).split()
    if any(difflib.SequenceMatcher(None, a, b).ratio() >= 0.75 for a in pf for b in pn):
        return True
    etiquetas = ("NOMBR", "FIRMA", "REPRESENTANTE", "LEGAL", "TRABAJADOR", "CONTRATISTA")
    reales = [w for w in pf if len(w) >= 4 and not w.startswith(etiquetas)]
    return len(reales) >= 2   # un nombre escrito completo y distinto: sí se compara (puede ser otra persona)


def _cmp_nombre(d: Doc, a: str, b: str, campo: str, etiqueta: str):
    if not a or not b:
        return
    if normaliza(a) == normaliza(b):  # ignora tildes, ñ, mayúsculas y espacios dobles
        return
    if compacto(a) == compacto(b):
        d.add(ALERTA, campo, f"{etiqueta}: difiere solo en espacios ('{a}' vs '{b}').")
        return
    ta, tb = set(normaliza(a).split()), set(normaliza(b).split())
    if (ta <= tb or tb <= ta) and min(len(ta), len(tb)) >= 2:
        d.add(ALERTA, campo, f"{etiqueta}: uno de los nombres está incompleto: '{a}' vs '{b}'.")
        return
    if sorted(_palabras_nombre(a)) == sorted(_palabras_nombre(b)) and _palabras_nombre(a):
        d.add(ALERTA, campo, f"{etiqueta}: difiere solo en la ortografía (y/j, b/v, s/z, h...): '{a}' vs '{b}'.")
        return
    s = similitud(a, b)
    if s >= 0.85:
        extra = " (leído por OCR: confirmar visualmente)" if "OCR" in d.fuente else ""
        d.add(ALERTA, campo, f"{etiqueta}: casi igual ({s:.0%}): '{a}' vs '{b}'.{extra}")
    else:
        d.add(ERROR, campo, f"{etiqueta}: no coincide: '{a}' vs '{b}'.")


def cobertura_cargo(a: str, b: str) -> float:
    ta, tb = set(tokens_cargo(a)), set(tokens_cargo(b))
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / min(len(ta), len(tb))


def _cmp_cargo(d: Doc, a: str, b: str, campo: str, etiqueta: str, max_nivel: str = ERROR):
    if not a or not b:
        return
    sa, sb = set(tokens_cargo(a)), set(tokens_cargo(b))
    if sa == sb:
        return
    corto, largo = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
    inter = len(sa & sb)
    if corto <= largo and len(corto) >= 2 and len(corto) / len(largo) >= 0.6:
        return  # un cargo contiene al otro
    if inter >= 3 and inter / len(corto) >= 0.75:
        return  # casi la misma descripción (p. ej. "Ingeniero para actas" vs "Ingeniero / Arquitecto - actas")
    if inter or similitud(a, b) >= 0.75 or max_nivel == ALERTA:
        d.add(ALERTA, campo, f"{etiqueta}: cargo no idéntico: '{a}' vs '{b}'.")
    else:
        d.add(ERROR, campo, f"{etiqueta}: cargo distinto: '{a}' vs '{b}'.")


def _cmp_fecha(d: Doc, f1, f2, campo: str, txt1: str, txt2: str):
    if f1 is None or f2 is None:
        return
    if f1 != f2:
        delta = abs((f1 - f2).days)
        d.add(ERROR, campo, f"{txt1} {fmt(f1)} ≠ {txt2} {fmt(f2)} (diferencia de {delta} día(s)).")


def clave_asunto(texto: str) -> str:
    """Clasifica un asunto: LABORAL, OPS_I_III, OPS_IV_V, OPS (sin clase), CONTINUIDAD o ''."""
    c = compacto(texto)
    if "CONTINUIDAD" in c:
        return "CONTINUIDAD"
    if "PRESTACIONDESERVICIO" in c:
        m = re.search(r"RIESGOLABORAL([IVY]+)", c)
        if m:
            r = m[1].replace("Y", "")
            if r == "IVV":
                return "OPS_IV_V"
            if set(r) == {"I"} and len(r) == 6:
                return "OPS_I_III"
        return "OPS"
    if "CONTRATOLABORAL" in c:
        return "LABORAL"
    return ""


def clave_esperada(tipo: str, clase: str) -> str:
    if tipo == "LABORAL":
        return "LABORAL"
    if tipo == "OPS":
        if clase in ("I", "II", "III"):
            return "OPS_I_III"
        if clase in ("IV", "V"):
            return "OPS_IV_V"
    return ""


def _dias(a, b):
    return abs((a - b).days) if a and b else None


def elegir_fila(d: Doc, filas: list[dict], desde, hasta) -> dict:
    """Empareja el documento con la fila del FO-AC-83 de su mismo periodo / cargo / empresa."""
    def puntaje(x):
        s = 0.0
        for fd, fx in ((desde, x["ingreso"]), (hasta, x["retiro"])):
            dd = _dias(fd, fx)
            if dd is not None:
                s += 3 if dd == 0 else 1 if dd <= 3 else 0
        s += 2 * cobertura_cargo(d.cargo, x["cargo"])
        s += similitud(d.empresa, x["empresa"]) if d.empresa and x["empresa"] else 0
        return s
    return max(filas, key=puntaje)


def validar_documento(d: Doc, fo83: dict | None, config: dict):
    cfg = {**CONFIG_DEFECTO, **(config or {})}

    # --- Estructura del documento ---
    if not d.formato:
        d.add(ALERTA, "formato", "No se leyó el código del formato (FO-AC-101 / FO-AC-103); se infirió por el contenido.")
    for campo, etiqueta in (("nombre", "nombre"), ("cedula", "cédula"), ("cargo", "cargo"),
                            ("desde_raw", "fecha inicial del periodo"), ("hasta_raw", "fecha final del periodo"),
                            ("empresa", "empresa"), ("fecha_doc_raw", "fecha del documento")):
        if not getattr(d, campo):
            if campo == "fecha_doc_raw" and re.search(r"«\s*D[IÍ]A\s*»|«\s*MES\s*»|«\s*A[NÑ]O\s*»", d.texto or "", re.I):
                d.add(ERROR, campo, "La fecha del documento NO está diligenciada: quedó el texto de la plantilla "
                                    "(«DÍA» de «MES» del «AÑO»).")
                continue
            d.add(ERROR, campo, f"No se pudo leer {etiqueta} (ni en el texto del PDF ni por OCR).")
    if d.formato == "FO-AC-103" and not d.hasta_contrato_raw:
        d.add(ALERTA, "hasta_contrato", "No se pudo leer la fecha hasta la que estuvo en el contrato.")
    if d.frases_faltantes:
        d.add(ALERTA, "plantilla", "Texto de plantilla no encontrado: " + "; ".join(d.frases_faltantes) + ".")

    # --- Cédula ---
    if d.cedula_raw and not cedula_formato_ok(d.cedula_raw):
        d.add(ALERTA, "cedula", f"Cédula con formato irregular: '{d.cedula_raw}' (¿puntos mal ubicados?).")
    if d.formato == "FO-AC-103" and d.cedula_firma_raw and _es_cedula_representante(d, d.cedula_firma_raw):
        # En la carta de continuidad, lo leído como "cédula de la firma" era la del representante legal
        # (la del trabajador suele estar escrita a mano y no se pudo leer): no se compara.
        d.cedula_firma_raw = ""
    if d.cedula_firma_raw and d.cedula and solo_digitos(d.cedula_firma_raw) != d.cedula:
        if "OCR" in d.fuente:
            d.add(ALERTA, "cedula", f"Cédula del texto ({d.cedula_raw}) ≠ cédula leída en la firma "
                                    f"({d.cedula_firma_raw}). La firma se leyó por OCR (a veces es letra a mano): "
                                    f"confirmar visualmente.")
        else:
            d.add(ERROR, "cedula", f"Cédula del texto ({d.cedula_raw}) ≠ cédula del bloque de firma ({d.cedula_firma_raw}).")

    # --- Coherencia interna texto vs. bloque de firma ---
    if d.formato == "FO-AC-103" and d.nombre_firma and d.representante:
        rep_n = normaliza(re.split(r"IDENTIFICAD", fold(d.representante))[0])
        if rep_n and _es_representante(d.nombre_firma, rep_n) and not _es_representante(d.nombre, rep_n):
            d.nombre_firma = ""   # lo leído en la firma es el representante legal, no el trabajador
    if d.nombre and d.nombre_firma and "OCR" in d.fuente and not _firma_legible(d.nombre, d.nombre_firma):
        d.add(ALERTA, "nombre", f"El nombre de la firma no se pudo leer bien por OCR (probablemente escrito a mano): "
                                f"'{d.nombre_firma}'. Confirmar visualmente que corresponde a {d.nombre}.")
    else:
        n0 = len(d.hallazgos)
        _cmp_nombre(d, d.nombre, d.nombre_firma, "nombre", "Nombre del texto vs. firma")
        if "OCR" in d.fuente:  # la firma leída por OCR (a veces a mano) no basta para un ERROR: se confirma a ojo
            d.hallazgos[n0:] = [(ALERTA, h[1], h[2] + " La firma se leyó por OCR: confirmar visualmente.")
                                if h[0] == ERROR else h for h in d.hallazgos[n0:]]
    if d.formato == "FO-AC-101":
        _cmp_cargo(d, d.cargo, d.cargo_firma, "cargo", "Cargo del texto vs. firma", ALERTA)

    # --- Fechas ---
    desde, r1 = parse_fecha(d.desde_raw)
    hasta, r2 = parse_fecha(d.hasta_raw)
    # Si una fecha del periodo está claramente escrita mes/día/año ("3/20/2024"), la otra también lo está
    # ("2/7/2024" es 7 de febrero, no 2 de julio).
    if _es_mda(d.desde_raw) or _es_mda(d.hasta_raw):
        for c in ("desde_raw", "hasta_raw"):
            m = re.fullmatch(r"[\s_]*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})[\s_]*", getattr(d, c))
            if m and int(m[1]) <= 12 and int(m[2]) <= 12 and m[1] != m[2]:
                try:
                    f_ = dt.date(int(m[3]), int(m[1]), int(m[2]))
                except ValueError:
                    continue
                original = getattr(d, c).strip(" _")
                setattr(d, c, f"{f_.day:02d}/{f_.month:02d}/{f_.year}")
                if c == "desde_raw":
                    desde, r1 = f_, True
                else:
                    hasta, r2 = f_, True
                d.add(ALERTA, "fechas", f"La otra fecha del periodo está escrita mes/día/año, así que "
                                        f"'{original}' se leyó como {fmt(f_)}. Confirmar en el documento.")
    fdoc, _ = parse_fecha(d.fecha_doc_raw)
    hcon, r3 = parse_fecha(d.hasta_contrato_raw)
    mf = re.search(r"(\d{1,2})\D+?MES\s+(?:DE\s+)?([A-Z]+)\D+?(\d{4})", fold(d.fecha_firma_raw))
    ffirma = None
    if mf and mf[2] in MESES:
        try:
            ffirma = dt.date(int(mf[3]), MESES[mf[2]], int(mf[1]))
        except ValueError:
            pass
    for raw, ok, etq in ((d.desde_raw, r1, "inicial"), (d.hasta_raw, r2, "final"),
                         (d.hasta_contrato_raw, r3, "de desvinculación")):
        if raw and "[" in raw:
            d.add(ALERTA, "fechas", f"La fecha {etq} no tiene el año escrito ('{raw.split('[')[0].strip()}'); "
                                    f"se dedujo {raw.split('[')[1].strip(']')}. Confirmar en el documento.")
        elif raw and re.fullmatch(r"\s*_?\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*\d{4}\s*_?\s*", raw) and \
                int(re.findall(r"\d+", raw)[1]) > 12:
            d.add(ALERTA, "fechas", f"Fecha {etq} escrita como mes/día/año ('{raw}'): se leyó como {fmt(parse_fecha(raw)[0])}.")
        elif raw and not ok:
            d.add(ALERTA, "fechas", f"Fecha {etq} con formato irregular: '{raw}'.")
    if d.desde_raw and desde is None:
        d.add(ERROR, "fechas", f"Fecha inicial ilegible: '{d.desde_raw}'.")
    if d.hasta_raw and hasta is None:
        d.add(ERROR, "fechas", f"Fecha final ilegible: '{d.hasta_raw}'.")
    if desde and hasta and desde > hasta:
        d.add(ERROR, "fechas", f"Periodo invertido: {fmt(desde)} > {fmt(hasta)}.")
    if d.formato == "FO-AC-103" and hasta and hcon and hasta != hcon:
        d.add(ERROR, "fechas", f"Fecha de desvinculación ({fmt(hcon)}) ≠ fin del paz y salvo ({fmt(hasta)}).")
    if fdoc and ffirma and fdoc != ffirma:
        d.add(ALERTA, "fechas", f"Fecha del encabezado ({fmt(fdoc)}) ≠ fecha de firma en el texto ({fmt(ffirma)}).")
    if fdoc and hasta and fdoc < hasta:
        d.add(ALERTA, "fechas", f"Documento fechado ({fmt(fdoc)}) antes del fin del periodo ({fmt(hasta)}).")

    # --- Cruce con FO-AC-83 ---
    if not fo83:
        return
    todas = []
    ceds_doc = [c for c in (d.cedula, solo_digitos(d.cedula_firma_raw)) if c]
    for ced in ceds_doc:
        todas = [x for x in fo83["filas"] if x["cedula"] == ced]
        if todas:
            break
    posibles = []
    if not todas and ceds_doc:
        # Cruce aproximado: >= 5 dígitos de la cédula en común y al menos un nombre o apellido parecido
        posibles = buscar_aproximada(d, ceds_doc, fo83["filas"])
        if posibles and (len(posibles) == 1 or (posibles[0]["ds"], posibles[0]["hits"]) >
                         (posibles[1]["ds"], posibles[1]["hits"])):
            a = posibles[0]
            todas = [x for x in fo83["filas"] if x["cedula"] == a["cedula"]]
            d.add(ERROR, "cedula", f"Cédula del certificado ({a['ced_doc']}) no está en el FO-AC-83: cruce aproximado con "
                                   f"{a['cedula']} ({a['nombre']}), coinciden {a['ds']} dígitos y {a['hits']} "
                                   "nombre(s)/apellido(s). Verificar que sea la misma persona y corregir la cédula.")
            posibles = []
    if not todas and not posibles and d.nombre:
        # Último recurso: la cédula del texto no está en el FO-AC-83 (ni parecida), pero el nombre completo sí,
        # en una sola persona del retiro (p. ej. el trabajador escribió otra cédula en el texto y la correcta
        # a mano en la firma, que el OCR no alcanza a leer). Se cruza por nombre y se reporta la diferencia.
        base_n = [x for x in fo83["filas"] if "RETIRO" in x["seccion"]] if cfg["solo_retiro"] else fo83["filas"]
        nn = len(_palabras_nombre(d.nombre))
        por_nombre = [x for x in base_n if nn >= 3 and coincidencias_nombre(d.nombre, x["nombre"]) >= min(nn, 4)
                      and coincidencias_nombre(x["nombre"], d.nombre) >= 3]
        if len({x["cedula"] for x in por_nombre}) == 1:
            x = por_nombre[0]
            todas = [y for y in fo83["filas"] if y["cedula"] == x["cedula"]]
            d.add(ERROR, "cedula", f"La cédula escrita en el certificado ({d.cedula_raw or '?'}) NO coincide con la del "
                                   f"FO-AC-83 ({x['cedula']}), pero el nombre sí: {x['nombre']}. Se cruzó por nombre; "
                                   "revisar la firma del documento y corregir la cédula donde esté mal.")
    filas = [x for x in todas if "RETIRO" in x["seccion"]] if cfg["solo_retiro"] else todas
    universo = [x for x in fo83["filas"] if "RETIRO" in x["seccion"]] if cfg["solo_retiro"] else fo83["filas"]
    if todas and not filas:
        ingresos = [x["ingreso"] for x in todas if x["ingreso"]]
        if hasta and ingresos and all(hasta < i for i in ingresos):
            d.add(ALERTA, "seccion", f"Certificado de un periodo anterior (hasta {fmt(hasta)}): en este FO-AC-83 la persona "
                                     f"figura en '{todas[0]['seccion']}' con ingreso {fmt(min(ingresos))} (reingreso). "
                                     "Ese periodo no está en RETIRO de este reporte, por eso no se compara.")
        else:
            d.add(ERROR, "seccion", f"La cédula figura en el FO-AC-83 en la sección '{todas[0]['seccion']}', "
                                    "no en RETIRO ni INGRESO-RETIRO: no hay retiro reportado.")
        return
    if not filas:
        cercanas = [x for x in universo if d.cedula and distancia_1(x["cedula"], d.cedula)]
        if posibles:
            d.add(ERROR, "falta_excel", f"Cédula {d.cedula or '?'} no está en el FO-AC-83 y hay varias personas parecidas "
                                        "(no se pudo elegir una): " + "; ".join(f"{c['cedula']} {c['nombre']}" for c in posibles[:3]))
        elif cercanas:
            c = cercanas[0]
            d.add(ERROR, "falta_excel", f"Cédula {d.cedula} no está en el FO-AC-83; casi igual a {c['cedula']} "
                                        f"({c['nombre']}): ¿error de digitación u OCR?")
        else:
            por_nombre = [x for x in universo if d.nombre and coincidencias_nombre(d.nombre, x["nombre"]) >= 3]
            pista = ""
            if len(por_nombre) == 1:
                x = por_nombre[0]
                pista = (f" Por el nombre podría ser {x['nombre']} (cédula {x['cedula']} en el FO-AC-83): "
                         "verificar cuál cédula es la correcta.")
            d.add(ERROR, "falta_excel", f"Cédula {d.cedula or '?'} no aparece en el FO-AC-83"
                                        + (" (secciones RETIRO / INGRESO-RETIRO)." if cfg["solo_retiro"] else ".") + pista)
        return

    x = elegir_fila(d, filas, desde, hasta)
    d.excel = x
    d.candidatas = len(filas)

    mejor_nombre = max([n for n in (d.nombre, d.nombre_firma) if n] or [""], key=lambda n: similitud(n, x["nombre"]))
    _cmp_nombre(d, mejor_nombre, x["nombre"], "nombre", "Nombre vs. FO-AC-83")
    _cmp_cargo(d, d.cargo, x["cargo"], "cargo", "Cargo vs. FO-AC-83")
    _cmp_fecha(d, desde, x["ingreso"], "inicio", "Inicio del paz y salvo", "fecha de ingreso FO-AC-83")
    if x["retiro"] is None:
        d.add(ERROR, "fin", "El FO-AC-83 no reporta fecha de retiro para esta persona.")
    else:
        _cmp_fecha(d, hasta, x["retiro"], "fin", "Fin del paz y salvo", "fecha de retiro FO-AC-83")

    # Empresa
    if d.empresa and x["empresa"]:
        te = set(tokens_cargo(d.empresa)) - {"SA", "SAS", "S"}
        tx = set(tokens_cargo(x["empresa"]))
        if te and not te <= tx and similitud(d.empresa, x["empresa"]) < 0.6:
            d.add(ALERTA, "empresa", f"Empresa del documento '{d.empresa}' ≠ empresa del FO-AC-83 '{x['empresa']}'.")

    # Tipo de vinculación + clase de riesgo -> asunto del formato
    tipo, clase = x["tipo"], x["clase"]
    if d.formato == "FO-AC-103":
        # La carta de continuidad la puede presentar cualquier vinculación y clase de riesgo.
        if clave_asunto(d.asunto) != "CONTINUIDAD":
            d.add(ALERTA, "asunto", f"Asunto no reconocido como continuidad laboral: '{d.asunto}'.")
        return
    esperada = clave_esperada(tipo, clase)
    if not esperada:
        d.add(ALERTA, "vinculacion", f"Vinculación '{tipo}' / riesgo '{clase}' sin regla de asunto: revisión manual.")
        return
    if not d.asunto:
        d.add(ERROR, "asunto", "No se pudo leer el asunto del paz y salvo.")
        return
    esp_txt = cfg["asuntos"][esperada]
    if compacto(d.asunto) == compacto(esp_txt):
        return
    if clave_asunto(d.asunto) == esperada:
        d.add(ALERTA, "asunto", f"Asunto correcto en contenido pero con diferencias de texto: '{d.asunto}'"
                                + (" (OCR: confirmar visualmente)." if "OCR" in d.fuente else "."))
    else:
        d.add(ERROR, "asunto", f"Asunto no corresponde: FO-AC-83 dice {tipo}, riesgo clase {clase}; se esperaba "
                               f"'{esp_txt}' y el documento dice '{d.asunto}'.")


def revisar(pdfs: list[tuple[str, bytes]], fo83: dict | None, config: dict | None = None,
            ocr: bool = True) -> dict:
    """pdfs: lista de (nombre_archivo, bytes). Devuelve docs + cruces globales."""
    cfg = {**CONFIG_DEFECTO, **(config or {})}
    docs: list[Doc] = []
    avisos: list[str] = []
    otros: list[dict] = []
    for nombre_arch, data in pdfs:
        try:
            paginas = leer_pdf(data, cfg, ocr, avisos_arch := [])
        except Exception as e:  # noqa: BLE001
            avisos.append(f"{nombre_arch}: no se pudo abrir el PDF ({e}).")
            continue
        avisos += [f"{nombre_arch} · {a}" for a in avisos_arch]
        rend = Renderizador(data)
        for g_ in agrupar_documentos(paginas, cfg["una_pagina_por_documento"]):
            d = extraer_campos(g_["texto"], g_["paginas"], g_["fuentes"])
            d.archivo = nombre_arch
            if "ILEGIBLE" in d.fuente or not d.es_formato:
                resumen = re.sub(r"\s+", " ", d.texto)[:160]
                otros.append({"archivo": nombre_arch, "pagina": d.pagina, "lectura": d.fuente,
                              "motivo": ("ilegible" if "ILEGIBLE" in d.fuente else
                                         d.motivo_no_formato or "no es FO-AC-101/103"),
                              "inicio_del_texto": resumen})
                continue
            if d.formato == "FO-AC-103":  # el trabajador es el bloque de firma de abajo a la derecha
                aplicar_bloque_firma(d, data, g_["paginas"][-1] - 1, cfg, rend, ocr, avisos, nombre_arch)
            # Si el texto del PDF quedó incompleto (palabras pegadas, frases cortadas), se relee la página
            if "OCR" not in g_["fuentes"]:
                falta = campos_faltantes(d)
                for modo in ("xtol", "ocr"):
                    if not falta or (modo == "ocr" and not ocr):
                        break
                    try:
                        txt = "\n".join(texto_pagina(data, p - 1, modo, cfg, rend) for p in g_["paginas"])
                    except Exception as e:  # noqa: BLE001
                        avisos.append(f"{nombre_arch} · pág. {d.pagina}: no se pudo releer ({modo}): {e}")
                        break
                    d2 = extraer_campos(txt, g_["paginas"], {"OCR" if modo == "ocr" else "texto"})
                    fusionar(d, d2)
                    if modo == "ocr":
                        d.fuente = "texto+OCR"
                    falta = campos_faltantes(d)
            docs.append(d)
    # duplicados exactos (misma cédula, cargo y periodo)
    vistos: dict[tuple, list[Doc]] = {}
    for d in docs:
        if d.cedula:
            vistos.setdefault((d.cedula, normaliza(d.cargo), d.desde_raw, d.hasta_raw), []).append(d)
    for d in docs:
        validar_documento(d, fo83, cfg)
        rep = vistos.get((d.cedula, normaliza(d.cargo), d.desde_raw, d.hasta_raw), [])
        if d.cedula and len(rep) > 1:
            d.add(ALERTA, "duplicado", "Certificado repetido (misma cédula, cargo y periodo) en las páginas: "
                  + ", ".join(o.pagina for o in rep if o is not d))
    faltantes = []
    if fo83:
        usadas = {d.excel["fila_excel"] for d in docs if d.excel}
        con_doc = {d.cedula for d in docs}
        faltantes = [dict(x, entrego_otro=x["cedula"] in con_doc) for x in fo83["filas"]
                     if "RETIRO" in x["seccion"] and x["fila_excel"] not in usadas]
    return {"docs": docs, "faltantes": faltantes, "avisos": avisos, "otros": otros}


# ---------------------------------------------------------------------------
# Salida tabular
# ---------------------------------------------------------------------------
ICONO = {OK: "✅ OK", ALERTA: "⚠️ ALERTA", ERROR: "❌ ERROR"}


def tabla_resumen(docs: list[Doc]):
    import pandas as pd
    filas = []
    for d in docs:
        x = d.excel or {}
        filas.append({
            "Estado": ICONO[d.estado],
            "Pág.": d.pagina,
            "Formato": d.formato,
            "Lectura": d.fuente,
            "Nombre": d.nombre_firma or d.nombre,
            "Cédula": d.cedula_raw,
            "Cargo": d.cargo,
            "Desde": fmt(parse_fecha(d.desde_raw)[0]) or d.desde_raw,
            "Hasta": fmt(parse_fecha(d.hasta_raw)[0]) or d.hasta_raw,
            "Empresa": d.empresa,
            "Vinculación (FO-AC-83)": x.get("tipo", ""),
            "Asunto": clave_asunto(d.asunto),
            "Riesgo ARL": x.get("clase", ""),
            "Errores": sum(h[0] == ERROR for h in d.hallazgos),
            "Alertas": sum(h[0] == ALERTA for h in d.hallazgos),
        })
    return pd.DataFrame(filas)


def tabla_hallazgos(docs: list[Doc]):
    import pandas as pd
    filas = []
    for d in docs:
        for nivel, campo, msg in d.hallazgos:
            filas.append({"Pág.": d.pagina, "Nombre": d.nombre_firma or d.nombre, "Cédula": d.cedula_raw,
                          "Nivel": nivel, "Campo": campo, "Detalle": msg})
    return pd.DataFrame(filas, columns=["Pág.", "Nombre", "Cédula", "Nivel", "Campo", "Detalle"])


def tabla_faltantes(faltantes: list[dict]):
    import pandas as pd
    return pd.DataFrame([{
        "Nombre": x["nombre"], "Cédula": x["cedula"], "Cargo": x["cargo"], "Vinculación": x["tipo"],
        "Empresa": x["empresa"], "Ingreso": fmt(x["ingreso"]), "Retiro": fmt(x["retiro"]), "Riesgo ARL": x["clase"],
        "Entregó otro certificado": "Sí" if x.get("entrego_otro") else "No",
        "Fila FO-AC-83": x["fila_excel"]} for x in faltantes])


# ---------------------------------------------------------------------------
# Vista de comparación Excel vs. certificado
# ---------------------------------------------------------------------------
CATEGORIAS = {  # campo del hallazgo -> categoría que ve el usuario
    "nombre": "nombre", "cedula": "cedula", "cargo": "cargo", "inicio": "inicio", "fin": "fin",
    "empresa": "empresa", "asunto": "asunto", "vinculacion": "asunto",
    "falta_excel": "falta_excel", "seccion": "seccion",
}


def categoria(campo: str) -> str:
    return CATEGORIAS.get(campo, "otras")


def categorias_doc(d: Doc) -> list[str]:
    out = []
    for nivel, campo, _ in d.hallazgos:
        if nivel in (ALERTA, ERROR) and categoria(campo) not in out:
            out.append(categoria(campo))
    return out


def _nivel(d: Doc, cat: str):
    hs = [h for h in d.hallazgos if categoria(h[1]) == cat and h[0] in (ALERTA, ERROR)]
    if not hs:
        return OK, ""
    peor = max(hs, key=lambda h: ORDEN[h[0]])[0]
    return peor, " | ".join(h[2] for h in hs)


def filas_comparacion(d: Doc) -> list[dict]:
    """Filas Dato | Excel | Certificado | Resultado (solo si la cédula se cruzó con el FO-AC-83)."""
    x = d.excel
    if not x:
        return []
    desde, _ = parse_fecha(d.desde_raw)
    hasta, _ = parse_fecha(d.hasta_raw)
    pares = [
        ("nombre", "Nombre", x["nombre"], d.nombre_firma or d.nombre),
        ("cedula", "Cédula", x["cedula"], d.cedula_raw),
        ("cargo", "Cargo", x["cargo"], d.cargo),
        ("inicio", "Ingreso / inicio del periodo", fmt(x["ingreso"]), fmt(desde) or d.desde_raw),
        ("fin", "Retiro / fin del periodo", fmt(x["retiro"]) or "N/A", fmt(hasta) or d.hasta_raw),
        ("empresa", "Empresa", x["empresa"], d.empresa),
        ("asunto", "Vinculación y riesgo / asunto",
         f"{x['tipo']} · riesgo clase {x['clase'] or '?'}", d.asunto or d.formato),
    ]
    filas = []
    for cat, etq, ex, pdf in pares:
        nivel, nota = _nivel(d, cat)
        filas.append({"cat": cat, "dato": etq, "excel": ex, "pdf": pdf, "nivel": nivel, "nota": nota})
    return filas


def otras_observaciones(d: Doc) -> list[tuple]:
    return [h for h in d.hallazgos if h[0] in (ALERTA, ERROR) and categoria(h[1]) == "otras"]


def tabla_comparacion(docs: list[Doc]):
    import pandas as pd
    filas = []
    for d in docs:
        for f in filas_comparacion(d):
            filas.append({"Pág.": d.pagina, "Nombre": d.nombre_firma or d.nombre, "Cédula": d.cedula_raw,
                          "Dato": f["dato"], "FO-AC-83 (Excel)": f["excel"], "Certificado (PDF)": f["pdf"],
                          "Resultado": {OK: "Coincide", ALERTA: "Revisar", ERROR: "NO COINCIDE"}[f["nivel"]],
                          "Detalle": f["nota"]})
    return pd.DataFrame(filas, columns=["Pág.", "Nombre", "Cédula", "Dato", "FO-AC-83 (Excel)",
                                        "Certificado (PDF)", "Resultado", "Detalle"])


ETIQUETA_CAT = {
    "nombre": "Nombre", "cedula": "Cédula", "cargo": "Cargo", "inicio": "Fecha de ingreso",
    "fin": "Fecha de retiro", "empresa": "Empresa", "asunto": "Asunto / vinculación / riesgo",
    "seccion": "No está en RETIRO del FO-AC-83", "falta_excel": "No está en el FO-AC-83", "otras": "Otras observaciones",
}
ETIQUETA_OTRAS = {
    "fechas": "Fechas del certificado", "fecha_doc_raw": "Fecha del documento", "plantilla": "Texto del formato",
    "duplicado": "Certificado repetido", "formato": "Código del formato", "hasta_contrato": "Fecha de desvinculación",
    "desde_raw": "Fecha inicial no leída", "hasta_raw": "Fecha final no leída",
}
QUE_HACER = {
    ERROR: "Los datos NO son iguales: corregir el certificado o el FO-AC-83.",
    ALERTA: "Mirar el documento: puede ser un error de escritura o de lectura del escáner.",
}


def explicacion_simple(cat: str, nivel: str, excel, pdf) -> str:
    """La diferencia dicha en palabras sencillas (para la ficha del informe)."""
    malo = nivel == ERROR
    t = {
        "nombre": "El nombre NO es el mismo." if malo else "El nombre está escrito un poco distinto: confirmar que es la misma persona.",
        "cedula": "La cédula NO es la misma: hay que corregirla." if malo else "Revisar la cédula en el documento.",
        "cargo": "El cargo NO es el mismo." if malo else "El cargo está escrito distinto (puede ser el mismo con otras palabras).",
        "inicio": f"El certificado empieza el {pdf or '—'}, pero en el FO-AC-83 la persona ingresó el {excel or '—'}.",
        "fin": f"El certificado termina el {pdf or '—'}, pero en el FO-AC-83 la persona se retiró el {excel or '—'}.",
        "empresa": "Es otra empresa." if malo else "La empresa está escrita distinto (nombre completo o sigla): confirmar que es la misma.",
        "asunto": "El asunto del certificado no corresponde a la vinculación o al riesgo del FO-AC-83.",
    }
    return t.get(cat, QUE_HACER.get(nivel, ""))


def exportar_excel(res: dict) -> bytes:
    """Informe en Excel pensado para que cualquiera lo entienda:
    1. Resumen: cuántos certificados hay, cuántos están bien y la lista de personas con lo que no coincide.
    2. Por persona: una ficha por certificado, igual que en la aplicación (Dato | FO-AC-83 | Certificado | Resultado).
    3. Retiros sin soporte, 4. Páginas no revisadas y 5. Detalle técnico (todos los hallazgos)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    docs, faltantes, otros = res["docs"], res.get("faltantes", []), res.get("otros", [])
    ROJO, AMAR, VERDE, GRIS, AZUL = "FDECEA", "FFF4D6", "E6F4EA", "F3F4F6", "1F3B5C"
    T_ROJO, T_AMAR, T_VERDE = "B42318", "9A6700", "1E7B45"
    fill = lambda c: PatternFill("solid", fgColor=c)
    fino = Side(style="thin", color="D0D5DD")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    envolver = Alignment(wrap_text=True, vertical="top")
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ESTADO = {ERROR: ("✖ CON ERRORES", ROJO, T_ROJO), ALERTA: ("⚠ REVISAR", AMAR, T_AMAR), OK: ("✔ TODO BIEN", VERDE, T_VERDE)}
    RES_FILA = {ERROR: ("✖ NO COINCIDE", ROJO, T_ROJO), ALERTA: ("⚠ REVISAR", AMAR, T_AMAR), OK: ("✔ Coincide", None, T_VERDE)}

    def celda(ws, r, c, v, negrita=False, color=None, texto=None, tam=11, al=envolver, b=True):
        x = ws.cell(row=r, column=c, value=v)
        x.font = Font(bold=negrita, color=texto, size=tam)
        x.alignment = al
        if color:
            x.fill = fill(color)
        if b:
            x.border = borde
        return x

    def titulo(ws, r, texto, ncol, color=AZUL, tam=14):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        x = ws.cell(row=r, column=1, value=texto)
        x.font = Font(bold=True, size=tam, color="FFFFFF")
        x.fill = fill(color)
        x.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 26 if tam >= 14 else 20

    def nota(ws, r, texto, ncol, color=None, negrita=False, alto=None, texto_color=None):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        x = ws.cell(row=r, column=1, value=texto)
        x.font = Font(bold=negrita, color=texto_color)
        x.alignment = envolver
        if color:
            x.fill = fill(color)
        if alto:
            ws.row_dimensions[r].height = alto

    nombre_de = lambda d: (d.excel or {}).get("nombre") or d.nombre or d.nombre_firma or "(nombre no leído)"

    def que_falla(d, nivel):
        out = []
        for h in d.hallazgos:
            e = ETIQUETA_CAT.get(categoria(h[1]), h[1]) if categoria(h[1]) != "otras" else ETIQUETA_OTRAS.get(h[1], "Otras observaciones")
            if h[0] == nivel and e not in out:
                out.append(e)
        return ", ".join(out)
    orden = {ERROR: 0, ALERTA: 1, OK: 2}
    docs_ord = sorted(docs, key=lambda d: (orden[d.estado], d.archivo, int(re.sub(r"\D.*", "", d.pagina) or 0)))
    wb = Workbook()

    # ------------------------------------------------------------------ 2. Por persona (se arma primero para los enlaces)
    wp = wb.active
    wp.title = "Por persona"
    for col, w_ in zip("ABCDE", (30, 34, 34, 18, 60)):
        wp.column_dimensions[col].width = w_
    titulo(wp, 1, "FICHA POR PERSONA — FO-AC-83 (Excel) contra el paz y salvo / carta de continuidad (PDF)", 5)
    nota(wp, 2, "Cada recuadro es un certificado. Las filas en ROJO no coinciden; en AMARILLO hay que mirar el documento; "
                "en blanco con ✔ coinciden. Primero van los que tienen errores.", 5, GRIS, alto=32)
    r = 4
    fila_ficha = {}
    for n, d in enumerate(docs_ord, 1):
        etq, c_fondo, c_txt = ESTADO[d.estado]
        fila_ficha[id(d)] = r
        wp.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        celda(wp, r, 1, f"{n}. 👤 {nombre_de(d)}", negrita=True, tam=13, color=c_fondo, texto="000000")
        celda(wp, r, 4, etq, negrita=True, color=c_fondo, texto=c_txt, al=centro)
        e_, a_ = que_falla(d, ERROR), que_falla(d, ALERTA)
        resumen = "\n".join(x for x in ((f"✖ No coincide: {e_}" if e_ else ""), (f"⚠ Revisar: {a_}" if a_ else "")) if x)
        celda(wp, r, 5, resumen or "✔ Todo coincide", negrita=True, color=c_fondo, texto=c_txt)
        wp.row_dimensions[r].height = 34 if (e_ and a_) else 24
        r += 1
        info = (f"C.C. {d.cedula_raw or '—'}   ·   Archivo: {d.archivo}   ·   Página: {d.pagina}   ·   "
                f"Formato: {d.formato or '—'}   ·   Leído: {'con escáner (OCR)' if 'OCR' in d.fuente else 'texto del PDF'}")
        if d.excel:
            info += f"   ·   Fila {d.excel['fila_excel']} del FO-AC-83"
        nota(wp, r, info, 5, texto_color="475467")
        r += 1
        comp = filas_comparacion(d)
        if comp:
            for c, h in enumerate(("Dato", "📊 FO-AC-83 (Excel)", "📄 Certificado (PDF)", "Resultado", "Explicación"), 1):
                celda(wp, r, c, h, negrita=True, color="E4E7EC")
            r += 1
            for f in comp:
                etq_r, fondo, txt = RES_FILA[f["nivel"]]
                expl = ""
                if f["nivel"] != OK:
                    expl = explicacion_simple(f["cat"], f["nivel"], f["excel"], f["pdf"])
                    if f["nota"]:
                        expl += "\nDetalle: " + f["nota"]
                vals = (f["dato"], f["excel"] or "—", f["pdf"] or "—", etq_r, expl)
                for c, v_ in enumerate(vals, 1):
                    celda(wp, r, c, v_, negrita=(c in (1, 4)), color=fondo, texto=txt if c == 4 else None)
                r += 1
        else:
            msgs = " ".join(h[2] for h in d.hallazgos if categoria(h[1]) in ("falta_excel", "seccion"))
            nota(wp, r, "⚠ No se pudo comparar con el FO-AC-83. " + msgs, 5, ROJO, negrita=True, alto=46, texto_color=T_ROJO)
            r += 1
        for h in otras_observaciones(d):
            nota(wp, r, ("✖ " if h[0] == ERROR else "⚠ ") + h[2], 5, ROJO if h[0] == ERROR else AMAR,
                 texto_color=T_ROJO if h[0] == ERROR else T_AMAR, alto=max(18, 15 * (1 + len(h[2]) // 120)))
            r += 1
        r += 1  # espacio entre fichas
    wp.freeze_panes = "A4"

    # ------------------------------------------------------------------ 1. Resumen
    ws = wb.create_sheet("Resumen", 0)
    for col, w_ in zip("ABCDEFGH", (6, 16, 36, 16, 22, 8, 40, 40)):
        ws.column_dimensions[col].width = w_
    titulo(ws, 1, "REVISIÓN DE PAZ Y SALVOS (FO-AC-101) Y CARTAS DE CONTINUIDAD (FO-AC-103) vs. FO-AC-83", 8, tam=15)
    con_obs = [d for d in docs if d.estado != OK]
    n_err = sum(d.estado == ERROR for d in docs)
    n_al = sum(d.estado == ALERTA for d in docs)
    if not con_obs and not faltantes:
        nota(ws, 2, "✅ SIN OBSERVACIONES: todo coincide con el FO-AC-83.", 8, VERDE, True, 28, T_VERDE)
    else:
        nota(ws, 2, f"🚨 Hay {len(con_obs)} certificado(s) con observaciones"
                    + (f" y {len(faltantes)} persona(s) retirada(s) sin paz y salvo." if faltantes else "."),
             8, ROJO, True, 28, T_ROJO)
    kp = [("Certificados revisados", len(docs), GRIS, "000000"), ("✔ Todo bien", len(docs) - len(con_obs), VERDE, T_VERDE),
          ("✖ Con errores", n_err, ROJO, T_ROJO), ("⚠ Para revisar", n_al, AMAR, T_AMAR),
          ("Retiros sin paz y salvo", len(faltantes), ROJO if faltantes else VERDE, T_ROJO if faltantes else T_VERDE),
          ("Páginas que no son certificados", len(otros), GRIS, "000000")]
    r = 4
    for etq, n_, fondo, txt in kp:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        celda(ws, r, 2, etq, negrita=True, color=fondo)
        celda(ws, r, 3, None, color=fondo)
        celda(ws, r, 4, n_, negrita=True, tam=14, color=fondo, texto=txt, al=centro)
        r += 1
    r += 1
    conteo = {}
    for d in docs:
        for nivel in (ERROR, ALERTA):
            for e in filter(None, que_falla(d, nivel).split(", ")):
                conteo.setdefault(e, {ERROR: 0, ALERTA: 0})[nivel] += 1
    if conteo:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        celda(ws, r, 2, "¿Qué se encontró?  (número de certificados)", negrita=True, color="E4E7EC")
        celda(ws, r, 3, None, color="E4E7EC")
        celda(ws, r, 4, "✖ No coincide", negrita=True, color=ROJO, texto=T_ROJO, al=centro)
        celda(ws, r, 5, "⚠ Revisar", negrita=True, color=AMAR, texto=T_AMAR, al=centro)
        r += 1
        for e, n_ in sorted(conteo.items(), key=lambda kv: (-kv[1][ERROR], -kv[1][ALERTA])):
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            celda(ws, r, 2, e)
            celda(ws, r, 3, None)
            celda(ws, r, 4, n_[ERROR] or "", al=centro, texto=T_ROJO, negrita=True)
            celda(ws, r, 5, n_[ALERTA] or "", al=centro, texto=T_AMAR)
            r += 1
        r += 1
    nota(ws, r, "CÓMO LEER ESTE INFORME:  ✖ rojo = los datos NO son iguales (hay que corregir).   ⚠ amarillo = mirar el "
                "documento (puede ser letra a mano o un error del escáner).   ✔ verde = todo coincide.   "
                "Haga clic en el nombre de una persona para ir a su ficha.", 8, GRIS, alto=46)
    r += 2
    titulo(ws, r, "LISTA DE PERSONAS (primero las que tienen errores)", 8, tam=12)
    r += 1
    for c, h in enumerate(("#", "Resultado", "Persona", "Cédula", "Archivo · página", "Errores",
                           "✖ NO coincide (corregir)", "⚠ Revisar en el documento"), 1):
        celda(ws, r, c, h, negrita=True, color="E4E7EC", al=centro)
    ws.freeze_panes = ws.cell(row=r + 1, column=1)
    r += 1
    for n, d in enumerate(docs_ord, 1):
        etq, fondo, txt = ESTADO[d.estado]
        vals = (n, etq, nombre_de(d), d.cedula_raw or "—", f"{d.archivo} · pág. {d.pagina}",
                sum(h[0] == ERROR for h in d.hallazgos), que_falla(d, ERROR) or "—", que_falla(d, ALERTA) or "—")
        for c, v_ in enumerate(vals, 1):
            col_f = fondo if c == 2 else (ROJO if c == 7 and v_ != "—" else AMAR if c == 8 and v_ != "—" else None)
            x = celda(ws, r, c, v_, negrita=(c == 2), color=col_f,
                      texto=txt if c == 2 else None, al=centro if c in (1, 2, 6) else envolver)
        x = ws.cell(row=r, column=3)
        x.hyperlink = f"#'Por persona'!A{fila_ficha[id(d)]}"
        x.font = Font(color="1D4ED8", underline="single")
        r += 1

    # ------------------------------------------------------------------ 3. Retiros sin soporte
    wf = wb.create_sheet("Retiros sin paz y salvo")
    for col, w_ in zip("ABCDEFGHI", (34, 14, 30, 16, 22, 12, 12, 12, 40)):
        wf.column_dimensions[col].width = w_
    titulo(wf, 1, "PERSONAS RETIRADAS EN EL FO-AC-83 QUE NO TIENEN PAZ Y SALVO NI CARTA DE CONTINUIDAD", 9)
    nota(wf, 2, "Estas personas aparecen en RETIRO del FO-AC-83, pero en los PDF no hay un certificado que corresponda. "
                "Hay que pedir el documento al contratista.", 9, GRIS, alto=32)
    if not faltantes:
        nota(wf, 4, "✅ Todas las personas retiradas tienen su certificado.", 9, VERDE, True, 24, T_VERDE)
    else:
        for c, h in enumerate(("Persona", "Cédula", "Cargo", "Vinculación", "Empresa", "Ingreso", "Retiro",
                               "Fila FO-AC-83", "Observación"), 1):
            celda(wf, 4, c, h, negrita=True, color="E4E7EC", al=centro)
        for i, x in enumerate(faltantes, 5):
            obs = ("Entregó otro certificado, pero ninguno corresponde a este cargo / periodo."
                   if x.get("entrego_otro") else "No se encontró ningún certificado de esta persona.")
            for c, v_ in enumerate((x["nombre"], x["cedula"], x["cargo"], x["tipo"], x["empresa"], fmt(x["ingreso"]),
                                    fmt(x["retiro"]), x["fila_excel"], obs), 1):
                celda(wf, i, c, v_, negrita=(c == 1), color=ROJO if c in (1, 9) else None)

    # ------------------------------------------------------------------ 4. Páginas no revisadas
    wo = wb.create_sheet("Páginas no revisadas")
    for col, w_ in zip("ABCDE", (22, 8, 14, 50, 70)):
        wo.column_dimensions[col].width = w_
    titulo(wo, 1, "PÁGINAS QUE NO SON PAZ Y SALVO NI CARTA DE CONTINUIDAD (no se revisaron)", 5)
    nota(wo, 2, "Portadas e instructivos del formato, volantes de liquidación, cartas y otros soportes. "
                "Si alguna en realidad es un certificado, revísela a mano.", 5, GRIS, alto=32)
    for c, h in enumerate(("Archivo", "Página", "Leído", "Por qué no se revisó", "Cómo empieza el texto"), 1):
        celda(wo, 4, c, h, negrita=True, color="E4E7EC", al=centro)
    for i, o in enumerate(otros, 5):
        for c, v_ in enumerate((o.get("archivo"), o.get("pagina"), o.get("lectura"), o.get("motivo"),
                                (o.get("inicio_del_texto") or "")[:200]), 1):
            celda(wo, i, c, v_, color=AMAR if c == 4 and "NO se entrega" in str(v_) else None)

    # ------------------------------------------------------------------ 5. Detalle técnico
    wt = wb.create_sheet("Detalle técnico")
    cols = ("Archivo", "Pág.", "Persona", "Cédula", "Nivel", "Campo", "Detalle")
    for col, w_ in zip("ABCDEFG", (20, 7, 32, 16, 10, 16, 100)):
        wt.column_dimensions[col].width = w_
    for c, h in enumerate(cols, 1):
        celda(wt, 1, c, h, negrita=True, color="E4E7EC", al=centro)
    r = 2
    for d in docs_ord:
        for nivel, campo, msg in d.hallazgos:
            color = ROJO if nivel == ERROR else AMAR if nivel == ALERTA else None
            for c, v_ in enumerate((d.archivo, d.pagina, nombre_de(d), d.cedula_raw, nivel, campo, msg), 1):
                celda(wt, r, c, v_, color=color if c == 5 else None)
            r += 1
    wt.auto_filter.ref = f"A1:G{max(r - 1, 1)}"
    wt.freeze_panes = "A2"

    for hoja in wb.worksheets:
        hoja.page_setup.orientation = "landscape"
        hoja.page_setup.fitToWidth = 1
        hoja.page_setup.fitToHeight = 0
        hoja.sheet_properties.pageSetUpPr.fitToPage = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    import sys
    pdf_path, xlsx_path = sys.argv[1], (sys.argv[2] if len(sys.argv) > 2 else None)
    fo = cargar_fo83(xlsx_path) if xlsx_path else None
    res = revisar([(os.path.basename(pdf_path), open(pdf_path, "rb").read())], fo)
    import pandas as pd
    pd.set_option("display.width", 250, "display.max_colwidth", 60, "display.max_columns", 30)
    print(tabla_resumen(res["docs"]).to_string())
    print()
    print(tabla_hallazgos(res["docs"]).to_string())
    print("\nRetiros sin soporte:", len(res["faltantes"]))
