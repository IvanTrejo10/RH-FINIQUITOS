"""Convenios y finiquitos: cálculos, textos y PDF.

Este módulo no usa Streamlit, así que se puede probar o reutilizar por separado.
Todo el dinero se maneja en centavos enteros para no arrastrar errores de punto flotante.

Bloques:
    Catálogos ............ estados, razones sociales, percepciones y descuentos
    Dinero y fechas ...... lectura de montos, número con letra, fechas con letra
    Cálculo y validación . totales por días de convenio y revisión de la captura
    Textos ............... redacción de convenio, renuncia y constancia
    Maquetado y PDF ...... posiciones dentro de la hoja carta y archivo PDF
    Firma e historial .... limpieza de la imagen de la firma y libro de Excel
"""
from __future__ import annotations

import base64
import io
import json
import re
import unicodedata
import zipfile
import zlib
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps
from pypdf import PdfReader
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as lienzo_pdf

RECURSOS = Path(__file__).resolve().parent / "assets"

# ------------------------------- Catálogos -------------------------------

ESTADOS = [
    "AGUASCALIENTES", "BAJA CALIFORNIA", "BAJA CALIFORNIA SUR", "CAMPECHE", "CHIAPAS", "CHIHUAHUA",
    "CIUDAD DE MÉXICO", "COAHUILA DE ZARAGOZA", "COLIMA", "DURANGO", "ESTADO DE MÉXICO", "GUANAJUATO",
    "GUERRERO", "HIDALGO", "JALISCO", "MICHOACÁN DE OCAMPO", "MORELOS", "NAYARIT", "NUEVO LEÓN", "OAXACA",
    "PUEBLA", "QUERÉTARO", "QUINTANA ROO", "SAN LUIS POTOSÍ", "SINALOA", "SONORA", "TABASCO", "TAMAULIPAS",
    "TLAXCALA", "VERACRUZ DE IGNACIO DE LA LLAVE", "YUCATÁN", "ZACATECAS",
]

# Razones sociales que venían en el catálogo del Excel (sugerencias; se puede escribir otra).
EMPRESAS = [
    "MINERA SANTA IRENE S.A. DE C.V.", "MAZ SOLUCIONES ECONOMICAS S.A DE C.V.",
    "ADMINISTRACION EN RECURSOS IN S.A. DE C.V.", "ADMINISTRADORA DE FONDOS GVL. S.A. DE C.V.",
    "ADMINISTRATION OF INTEGRAL SERVICES OGI S. DE R.L. DE C.V.", "AGROINDUSTRIAS VELIZ S.A. DE C.V.",
    "AUDACIA EFICAZ S.A. DE C.V.", "CONSORTIUM LEGAL S.A. DE C.V.",
    "COORDINACION DE RECURSOS ADMINISTRATIVOS S.A. DE C.V.", "COORDINACION EN TRANSACCIONES S.A. DE C.V.",
    "COPERARY S.A. DE C.V.", "CRECIENDO CONTIGO S.A.", "DEJUR S.A DE C.V.", "DISCIPLINA EN RECURSOS S.A. DE C.V.",
    "DTO SOLUCIONES INTEGRALES, S.C.", "ESTRUCTURA LEGAL S.A. DE C.V.", "ESTABILIDAD EN NEGOCIO SA DE CV",
    "FIRST IN SOLUTIONS S.A.C.", "GESTION DE NEGOCIOS DEL NORTE S.A. DE C.V.", "GESTION DE NEGOCIOS IN S.A. DE C.V.",
    "GESTION OPERATIVA ORO S.A. DE C.V.", "GRAFICOS VECTORES Y LIENZOS, S.A. DE C.V.",
    "INNOVACIONES COMPUTACIONALES VL S.A. DE C.V.", "INNOVADORA DE PRESTIGIO, S.A. DE C.V.",
    "MAZ ESTRUCTURA OPERACIONAL SA DE CV", "MAZ RECURSOS ADMINISTRATIVOS S.A. DE C.V.",
    "MAZ SERVICIOS E INVERSIONES S.A DE C.V.", "MAZ TRANSACCIONES ADMINISTRATIVAS, S.A DE C.V.",
    "MINERA CAPPREPA S.A. DE C.V.", "POSICION LEGAL, S.A. DE C.V.",
    "PROTECCION Y MONITOREO DEL PACIFICO S.A. DE C.V.", "PROVEEDORA DE SOLVENCIAS S.A. DE C.V.",
    "RECURSOS ADMINISTRATIVOS C&F S.A. DE C.V.", "RECURSOS EN TRANSACCIONES S.A. DE C.V.", "RECURSOS IN S.A. DE C.V.",
    "RESPALDO EN VALOR S.A. DE C.V.", "SOLIDEZ CALIFICADA S.A. DE C.V.", "SOLUCIONES COMPARATIVAS S.A. DE C.V.",
    "SOLUCIONES ECONOMICAS C&F S.A. DE C.V.", "SOLUCIONES ECONOMICAS HI S.A. DE C.V.",
    "SOLVENCIAS PERSONALES EN EDUCACION, S.A. DE C.V.", "SOPORTE Y FIANZAS EN VALORES S.A. DE C.V.",
    "TEMPLANZA LEGAL S.A. DE C.V.", "TRANSPARENCIA OPERATIVA S.A. DE C.V.", "TURIMAZ RENT A CAR S.A. DE C.V.",
    "TENDENCIA ADMINISTRATIVA SA DE CV",
]

# Renglones del finiquito. "doc" es la etiqueta que se imprime en el convenio.
PERCEPCIONES = [
    {"k": "salarios", "nombre": "Salarios pendientes", "doc": "SALARIOS PENDIENTES", "dias": True},
    {"k": "descanso", "nombre": "Descanso y festivo laborado", "doc": "DESCANSO Y FESTIVO LABORADO", "dias": True},
    {"k": "aguinaldo", "nombre": "Aguinaldo proporcional", "doc": "AGUINALDO PROPORCIONAL", "dias": True, "periodo": True, "obligatorio": True},
    {"k": "vacaciones", "nombre": "Vacaciones proporcionales", "doc": "VACACIONES PROPORCIONALES", "dias": True, "periodo": True, "obligatorio": True},
    {"k": "prima", "nombre": "Prima vacacional proporcional", "doc": "PRIMA VACACIONAL PROPORCIONAL", "dias": True, "periodo": True, "obligatorio": True},
    {"k": "fondoAhorro", "nombre": "Liquidación de fondo de ahorro", "doc": "LIQUIDACION FONDO AHORRO"},
    {"k": "vacPend", "nombre": "Vacaciones pendientes", "doc": "VACACIONES PENDIENTES", "dias": True},
    {"k": "fondoRetiro", "nombre": "Fondo de retiro", "doc": "FONDO DE RETIRO"},
    {"k": "otras", "nombre": "Otras prestaciones exentas", "doc": "OTRAS PRESTACIONES EXCENTAS", "negrita": True},
]

DESCUENTOS = [
    {"k": "faltante", "nombre": "Faltante", "doc": "DESCUENTO POR FALTANTE"},
    {"k": "diasNomina", "nombre": "Días pagados de nómina", "doc": "DESCUENTO POR DIAS PAGADOS DE NOMINA"},
    {"k": "cajaChica", "nombre": "Caja chica", "doc": "DESCUENTO POR CAJA CHICA"},
    {"k": "infonavit", "nombre": "Infonavit", "doc": "DESCUENTO POR INFONAVIT"},
    {"k": "prestamo", "nombre": "Préstamo", "doc": "DESCUENTO POR PRESTAMO"},
    {"k": "siniestro", "nombre": "Siniestro", "doc": "DESCUENTO POR SINIESTRO"},
    {"k": "viaticos", "nombre": "Viáticos", "doc": "DESCUENTO POR VIATICOS"},
    {"k": "otros", "nombre": "Otros descuentos", "doc": "OTROS DESCUENTOS"},
]

DIAS_CONVENIO = [0, 15, 30, 45]

# El servidor puede estar en otro huso horario: la hora del convenio se toma del estado.
ZONA_GENERAL = "America/Mexico_City"
ZONAS = {
    "BAJA CALIFORNIA": "America/Tijuana", "BAJA CALIFORNIA SUR": "America/Mazatlan", "SINALOA": "America/Mazatlan",
    "NAYARIT": "America/Mazatlan", "SONORA": "America/Hermosillo", "QUINTANA ROO": "America/Cancun",
}

# ------------------------------- Dinero ----------------------------------

_CON_COMAS = re.compile(r"^\d{1,3}(,\d{3})+(\.\d*)?$")
_SIMPLE = re.compile(r"^\d*\.?\d*$")


def a_centavos(v) -> int | None:
    """"$7,274.01" -> 727401 · vacío -> None · texto que no es un monto -> ValueError."""
    if v is None:
        return None
    s = re.sub(r"[$\s]", "", str(v))
    if s == "":
        return None
    if not (_CON_COMAS.match(s) or (_SIMPLE.match(s) and re.search(r"\d", s))):
        raise ValueError(f"No es un monto: {v!r}")
    try:
        n = Decimal(s.replace(",", ""))
    except InvalidOperation as error:
        raise ValueError(f"No es un monto: {v!r}") from error
    return int((n * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def centavos(v) -> int:
    """Igual que a_centavos, pero vacío o inválido cuenta como cero (para sumar)."""
    try:
        return a_centavos(v) or 0
    except ValueError:
        return 0


def dividir(a: int, b: int) -> int:
    """Redondeo "mitad hacia arriba" de a/b con enteros no negativos (como REDONDEAR de Excel)."""
    return (2 * a + b) // (2 * b)


def numero(c: int) -> str:
    """1100000 -> "11,000.00"."""
    signo = "-" if c < 0 else ""
    c = abs(int(c))
    return f"{signo}{c // 100:,}.{c % 100:02d}"


def pesos(c: int) -> str:
    return ("-$" if c < 0 else "$") + numero(abs(c))


def a_numero(v) -> float | None:
    """Días u otras cantidades: "7.270" -> 7.27 · vacío -> None · inválido -> ValueError."""
    if v is None:
        return None
    s = str(v).strip().replace(",", "")
    if s == "":
        return None
    if not (_SIMPLE.match(s) and re.search(r"\d", s)):
        raise ValueError(f"No es un número: {v!r}")
    return float(s)


# --------------------------- Número con letra ----------------------------
# Misma salida que la función NumLetras de la macro ("UN MIL", sin acentos).

_UNIDADES = ["UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE",
             "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN",
             "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO",
             "VEINTINUEVE"]
_DECENAS = ["DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
_CENTENAS = ["CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS",
             "OCHOCIENTOS", "NOVECIENTOS"]


def _bloque_letras(n: int) -> str:  # 1..999
    c, du = divmod(n, 100)
    d, u = divmod(du, 10)
    partes = []
    if c:
        partes.append("CIEN" if c == 1 and du == 0 else _CENTENAS[c - 1])
    if du:
        partes.append(_UNIDADES[du - 1] if du <= 29 else _DECENAS[d - 1] + (" Y " + _UNIDADES[u - 1] if u else ""))
    return " ".join(partes)


def num_letras(c: int) -> str:
    """1100000 (centavos) -> "ONCE MIL PESOS 00/100 MN"."""
    c = abs(int(c))
    entero, cc = divmod(c, 100)
    if entero >= 10 ** 9:  # fuera del alcance de la función original
        return numero(c) + " PESOS MN"
    u, m, mm = entero % 1000, (entero // 1000) % 1000, entero // 10 ** 6
    partes = []
    if mm:
        partes.append(_bloque_letras(mm) + (" MILLON" if mm == 1 else " MILLONES"))
    if m:
        partes.append(_bloque_letras(m) + " MIL")
    if u:
        partes.append(_bloque_letras(u))
    letras = " ".join(partes) or "CERO"
    if mm and not m and not u:
        letras += " DE"
    return f"{letras}{' PESO ' if entero == 1 else ' PESOS '}{cc:02d}/100 MN"


# -------------------------------- Fechas ---------------------------------

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]
DIAS_SEMANA = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


_FECHA_ESCRITA = re.compile(r"(\d{1,2})\s*[/\-. ]\s*(\d{1,2})\s*[/\-. ]\s*(\d{4}|\d{2})")  # 20/02/2026, 20-2-26
_FECHA_JUNTA = re.compile(r"(\d{2})(\d{2})(\d{4}|\d{2})")  # 20022026, 200226


def a_fecha(v) -> date | None:
    """Convierte a fecha lo capturado; devuelve None si está vacío o no es una fecha que exista.

    Acepta date, datetime, texto día/mes/año ("20/02/2026", "20-2-26", "20022026") y "2026-02-20".
    Un año de dos dígitos se toma como 20xx.
    """
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = limpio(v)
    m = _FECHA_ESCRITA.fullmatch(s) or _FECHA_JUNTA.fullmatch(s)
    try:
        if m:
            dia, mes, anio = (int(x) for x in m.groups())
            return date(anio + 2000 if anio < 100 else anio, mes, dia)
        return date.fromisoformat(s[:10])
    except ValueError:
        return None


def fecha_corta(f: date) -> str:
    return f"{f.day:02d}/{f.month:02d}/{f.year}"  # 20/02/2026


def fecha_texto(f: date) -> str:
    return f"{f.day:02d} de {MESES[f.month - 1]} de {f.year}"  # 20 de febrero de 2026


def fecha_con_dia(f: date) -> str:
    return f"{DIAS_SEMANA[f.weekday()]}, {fecha_texto(f)}"


def ahora_local(estado: str = "") -> datetime:
    """Fecha y hora actuales en el huso del estado (el servidor puede estar en otro huso)."""
    return datetime.now(ZoneInfo(ZONAS.get(may(estado), ZONA_GENERAL)))


def hora_local(estado: str = "", ahora: datetime | None = None) -> str:
    """Hora actual (HH:MM) en el huso del estado donde se firma."""
    momento = ahora or ahora_local(estado)
    return f"{momento.hour:02d}:{momento.minute:02d}"


# ------------------------------- Captura ---------------------------------

def limpio(v) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", "" if v is None else str(v))).strip()


def may(v) -> str:
    return limpio(v).upper()


def sin_acentos(v) -> str:
    """Mayúsculas sin acentos, para comparar textos (la Ñ se conserva)."""
    texto = unicodedata.normalize("NFD", may(v).replace("Ñ", "\x00"))
    return "".join(ch for ch in texto if not unicodedata.combining(ch)).replace("\x00", "Ñ")


def captura_vacia() -> dict:
    return {
        "empresa": "", "sucursal": "", "estado": "", "ciudad": "", "representante": "",
        "nombre": "", "idNomina": "", "sueldo": "", "ingreso": None, "baja": None, "hora": "", "rfc": "", "nss": "",
        "incluirHora": True, "incluirRfc": True, "incluirNss": True,
        "perc": {p["k"]: {"periodo": "", "dias": "", "monto": ""} for p in PERCEPCIONES},
        "desc": {d["k"]: "" for d in DESCUENTOS},
    }


def firmante_vacio() -> dict:
    return {"nombre": "", "area": "RECURSOS HUMANOS", "telefono": "", "correo": "", "direccion": "", "firma": None}


# -------------------------------- Cálculo --------------------------------

def calcular(s: dict) -> dict:
    sueldo = centavos(s.get("sueldo"))
    baja, ingreso = a_fecha(s.get("baja")), a_fecha(s.get("ingreso"))
    perc = []
    for p in PERCEPCIONES:
        r = (s.get("perc") or {}).get(p["k"]) or {}
        try:
            dias = a_numero(r.get("dias")) or 0
        except ValueError:
            dias = 0
        periodo = (limpio(r.get("periodo")) or (str(baja.year) if baja else "")) if p.get("periodo") else ""
        perc.append({"k": p["k"], "doc": p["doc"], "negrita": bool(p.get("negrita")),
                     "monto": centavos(r.get("monto")), "dias": dias, "periodo": periodo})
    desc = [{"k": d["k"], "doc": d["doc"], "monto": centavos((s.get("desc") or {}).get(d["k"]))} for d in DESCUENTOS]
    percepciones = sum(p["monto"] for p in perc)
    deducciones = sum(d["monto"] for d in desc)
    otras_base = next(p["monto"] for p in perc if p["k"] == "otras")
    opciones = {}
    for n in DIAS_CONVENIO:
        extra = dividir(sueldo * n, 30)  # sueldo diario × días de convenio
        opciones[n] = {
            "dias": n, "extra": extra,
            "otras": otras_base + extra,  # se suma a "otras prestaciones exentas"
            "percepciones": percepciones + extra,
            "deducciones": deducciones,
            "total": percepciones + extra - deducciones,
        }
    return {"sueldo": sueldo, "diario": dividir(sueldo, 30), "baja": baja, "ingreso": ingreso, "perc": perc,
            "desc": desc, "percepciones": percepciones, "deducciones": deducciones, "opciones": opciones}


# ------------------------------ Validación -------------------------------

_PIDE = {
    "convenio": ["empresa", "representante", "sucursal", "estado", "ciudad", "nombre", "ingreso", "baja", "sueldo"],
    "renuncia": ["empresa", "estado", "ciudad", "nombre", "baja"],
    "constancia": ["empresa", "estado", "ciudad", "nombre", "ingreso", "baja"],
}
_NOMBRES = {
    "empresa": "la empresa", "sucursal": "la sucursal", "estado": "el estado", "ciudad": "la ciudad",
    "representante": "el representante legal", "nombre": "el nombre del colaborador", "sueldo": "el sueldo mensual",
    "ingreso": "la fecha de ingreso", "baja": "la fecha de baja", "rfc": "el RFC", "nss": "el número de seguro social",
}


def es_monto(v) -> bool:
    """True si está vacío o es un monto bien escrito."""
    try:
        a_centavos(v)
        return True
    except ValueError:
        return False


def es_numero(v) -> bool:
    """True si está vacío o es un número (días) bien escrito."""
    try:
        a_numero(v)
        return True
    except ValueError:
        return False


def _es_hora(v) -> bool:
    return re.fullmatch(r"([01]\d|2[0-3]):[0-5]\d", limpio(v)) is not None


def _es_anio(v) -> bool:
    return re.fullmatch(r"\d{4}", limpio(v)) is not None


def _es_nss(v) -> bool:
    return re.fullmatch(r"\d{11}", limpio(v)) is not None


def _sueldo_en_rango(v) -> bool:
    return 100 <= centavos(v) <= 30_000_000  # de $1 a $300,000


def _anio_en_rango(f: date) -> bool:
    return 2008 <= f.year <= 2040


AVISO_MONTO = "Escribe solo el número, por ejemplo 1,234.56"


def mal_escritos(s: dict) -> dict[str, str]:
    """Campos que tienen algo escrito pero con una forma que no se entiende: {campo: aviso corto}.

    Aplica las mismas reglas que validar(). Sirve para avisar junto a cada campo en cuanto se escribe,
    sin esperar a pedir el documento. Los nombres de campo son los mismos que devuelve validar().
    """
    e: dict[str, str] = {}
    fechas = {k: a_fecha(s.get(k)) for k in ("ingreso", "baja")}
    for k, f in fechas.items():
        if limpio(s.get(k)) and not f:
            e[k] = "Escribe la fecha como 20/02/2026"
        elif f and not _anio_en_rango(f):
            e[k] = "Revisa el año"
    if "baja" not in e and fechas["ingreso"] and fechas["baja"] and fechas["baja"] < fechas["ingreso"]:
        e["baja"] = "Es anterior a la fecha de ingreso"
    if not es_monto(s.get("sueldo")):
        e["sueldo"] = AVISO_MONTO
    elif limpio(s.get("sueldo")) and not _sueldo_en_rango(s["sueldo"]):
        e["sueldo"] = "Debe estar entre $1 y $300,000"
    for p in PERCEPCIONES:
        r = (s.get("perc") or {}).get(p["k"]) or {}
        if not es_monto(r.get("monto")):
            e[f"perc.{p['k']}.monto"] = AVISO_MONTO
        if not es_numero(r.get("dias")):
            e[f"perc.{p['k']}.dias"] = "Solo el número, por ejemplo 7.5"
        if p.get("periodo") and limpio(r.get("periodo")) and not _es_anio(r.get("periodo")):
            e[f"perc.{p['k']}.periodo"] = "Año con 4 dígitos"
    for d in DESCUENTOS:
        if not es_monto((s.get("desc") or {}).get(d["k"])):
            e[f"desc.{d['k']}"] = AVISO_MONTO
    if s.get("incluirHora", True) and limpio(s.get("hora")) and not _es_hora(s.get("hora")):
        e["hora"] = "Escribe la hora como 09:30"
    if s.get("incluirRfc", True) and limpio(s.get("rfc")) and len(limpio(s.get("rfc"))) != 13:
        e["rfc"] = "El RFC debe tener 13 caracteres"
    if s.get("incluirNss", True) and limpio(s.get("nss")) and not _es_nss(s.get("nss")):
        e["nss"] = "Deben ser 11 dígitos, sin espacios ni guiones"
    return e


def validar(s: dict, firmante: dict | None, documento: str) -> list[tuple[str, str]]:
    """Devuelve [(campo, texto)] con lo que falta o está mal para el documento pedido."""
    e: list[tuple[str, str]] = []
    baja, ingreso = a_fecha(s.get("baja")), a_fecha(s.get("ingreso"))
    for k in _PIDE[documento]:
        if not limpio(s.get(k)):
            e.append((k, "Falta " + _NOMBRES[k]))
    for k, f in (("ingreso", ingreso), ("baja", baja)):
        if limpio(s.get(k)) and not f:
            e.append((k, "No es válida " + _NOMBRES[k]))
        if f and not _anio_en_rango(f):
            e.append((k, "Revisa el año de " + _NOMBRES[k]))
    if ingreso and baja and baja < ingreso:
        e.append(("baja", "La fecha de baja es anterior a la de ingreso"))

    if documento == "convenio":
        if limpio(s.get("sueldo")):
            if not es_monto(s["sueldo"]) or not _sueldo_en_rango(s["sueldo"]):
                e.append(("sueldo", "El sueldo mensual debe estar entre $1 y $300,000"))
        for p in PERCEPCIONES:
            r = (s.get("perc") or {}).get(p["k"]) or {}
            de = p["nombre"].lower()
            if not es_monto(r.get("monto")):
                e.append((f"perc.{p['k']}.monto", "Revisa el monto de " + de))
            elif p.get("obligatorio") and a_centavos(r.get("monto")) is None:
                e.append((f"perc.{p['k']}.monto", f"Falta el monto de {de} (si no aplica, escribe 0)"))
            if not es_numero(r.get("dias")):
                e.append((f"perc.{p['k']}.dias", "Revisa los días de " + de))
            if p.get("periodo") and limpio(r.get("periodo")) and not _es_anio(r.get("periodo")):
                e.append((f"perc.{p['k']}.periodo", f"El año de {de} debe tener 4 dígitos"))
        for d in DESCUENTOS:
            if not es_monto((s.get("desc") or {}).get(d["k"])):
                e.append((f"desc.{d['k']}", "Revisa el descuento de " + d["nombre"].lower()))
        if s.get("incluirHora", True) and limpio(s.get("hora")) and not _es_hora(s.get("hora")):
            e.append(("hora", "La hora debe ir como 09:30"))
        if not e and calcular(s)["opciones"][0]["total"] < 0:
            e.append(("desc.faltante", "Los descuentos son mayores que las percepciones"))

    if documento == "constancia":
        if s.get("incluirRfc", True):
            if not limpio(s.get("rfc")):
                e.append(("rfc", "Falta " + _NOMBRES["rfc"]))
            elif len(limpio(s.get("rfc"))) != 13:
                e.append(("rfc", "El RFC debe tener 13 caracteres"))
        if s.get("incluirNss", True):
            if not limpio(s.get("nss")):
                e.append(("nss", "Falta " + _NOMBRES["nss"]))
            elif not _es_nss(s.get("nss")):
                e.append(("nss", "El número de seguro social debe tener 11 dígitos"))
        if not limpio((firmante or {}).get("nombre")):
            e.append(("firmante.nombre", "Falta el nombre de quien firma la constancia"))
    return e


# --------------------------- Datos para imprimir -------------------------

def datos(s: dict, firmante: dict | None = None, ahora: datetime | None = None) -> dict:
    f = firmante or {}
    calc = calcular(s)
    inc_hora = bool(s.get("incluirHora", True))
    inc_rfc = bool(s.get("incluirRfc", True))
    inc_nss = bool(s.get("incluirNss", True))
    return {
        "empresa": may(s.get("empresa")), "sucursal": may(s.get("sucursal")), "estado": may(s.get("estado")),
        "ciudad": may(s.get("ciudad")), "representante": may(s.get("representante")), "nombre": may(s.get("nombre")),
        "rfc": may(s.get("rfc")) if inc_rfc else "",
        "nss": limpio(s.get("nss")) if inc_nss else "",
        "idNomina": limpio(s.get("idNomina")),
        "ingreso": calc["ingreso"], "baja": calc["baja"],
        "hora": (limpio(s.get("hora")) or hora_local(s.get("estado"), ahora)) if inc_hora else "",
        "calc": calc,
        "firmante": {
            "nombre": may(f.get("nombre")), "area": may(f.get("area")), "telefono": limpio(f.get("telefono")),
            "correo": limpio(f.get("correo")),
            "direccion": [r for r in (limpio(x) for x in str(f.get("direccion") or "").splitlines()) if r],
            "firma": f.get("firma"),  # {"png": bytes, "ancho": int, "alto": int} o None
        },
    }


# ------------------------- Textos de los documentos ----------------------
# Redacción tomada del Excel original. Para cambiar una cláusula, edítala aquí.

def lugar_fecha(d: dict) -> str:
    return f"{d['ciudad']}, {d['estado']} A {fecha_texto(d['baja'])}.".upper()


TITULO_CONVENIO = "CONVENIO ACTO"
CLAUSULAS = "C L A U S U L A S"
SEGUNDA = ("SEGUNDA. - El (la) Trabajador(a) manifiesta que el Patrón le ha cubierto hasta la fecha todas y cada una de "
           "las prestaciones legales y contractuales ordinarias y extraordinarias a que tuvo derecho derivadas de su "
           "relación laboral.")
CUARTA = [
    ("CUARTA.- El (la) ", False), ("Trabajador(a)", True),
    (" manifiesta que está conforme con la cantidad recibida y que durante la relación de trabajo laboró siempre sujeto "
     "a una jornada ordinaria diurna de 8 horas diarias comprendidas de lunes a sábado, descansando los domingos de "
     "cada semana, así mismo manifiesta que jamás laboró horas extras, manifestando que no laboró en sus días de "
     "descanso semanal, ni en sus días de descanso obligatorio, y que no se le adeudan salarios devengados o bien "
     "cualquier otra prestación, por lo que no tiene reclamación alguna que hacer por dichos conceptos.", False),
]
FUNDAMENTO = [
    ("Con fundamento en el segundo párrafo del artículo 33 de la Ley Federal del Trabajo, venimos ratificando el "
     "presente ", False),
    ("CONVENIO", True),
    ("; el cual solicitamos se apruebe para que surta los efectos legales procedentes.", False),
]
UNICO = ("ÚNICO. - Se nos tenga por ratificado y se apruebe el presente convenio, debiéndose elevar la categoría de "
         "laudo.")
EL_TRABAJADOR, EL_PATRON = "El TRABAJADOR", "El PATRÓN"


def encabezado_convenio(d: dict) -> list[str]:
    return ["CENTRO DE CONCILIACIÓN Y ARBITRAJE", "EN EL ESTADO DE " + d["estado"], "P R E S E N T E.-"]


def intro_convenio(d: dict) -> str:
    cuando = (f"siendo las {d['hora']} horas del día {fecha_texto(d['baja'])}"
              if d["hora"] else f"el día {fecha_texto(d['baja'])}")
    return (f"En la ciudad de {d['ciudad']}, {d['estado']}, {cuando}, ante las oficinas de la H. Junta Local "
            f"de Conciliación y Arbitraje del Estado de {d['estado']} comparecemos los suscritos C. {d['nombre']} "
            f"como “TRABAJADOR(A)” y por la otra parte la empresa {d['empresa']} como “PATRÓN”, representada en "
            f"este acto por el (la) C. {d['representante']} quienes de común acuerdo celebramos el presente "
            "CONVENIO reconociéndonos la personalidad con que comparecemos, con arreglo a las siguientes:")


def primera(d: dict) -> str:
    diario = d["calc"]["diario"]
    return ("PRIMERA. - El (la) Trabajador(a) y Patrón, con fundamento en la fracción primera del artículo 53 de la "
            "Ley Federal del Trabajo, acuerdan en dar por terminada la relación laboral que los unía, la cual inició a "
            f"partir del día {fecha_texto(d['ingreso'])} hasta el día {fecha_texto(d['baja'])} y en la cual el (la) "
            f"Trabajador(a) percibía un salario diario de {pesos(diario)} ({num_letras(diario)}).")


def tercera(d: dict, n: int) -> str:
    # 0 días se paga por transferencia; 15, 30 y 45 días en efectivo (así venía en el Excel).
    total = d["calc"]["opciones"][n]["total"]
    forma = "TRANFERENCIA BANCARIA" if n == 0 else "EFECTIVO"
    return ("TERCERA. - El Patrón, en este acto y a la firma del presente convenio, le hace entrega al (la) "
            f"Trabajador(a) la cantidad de {pesos(total)} ({num_letras(total)}), PAGO REALIZADO EN {forma} POR LOS "
            "CONCEPTOS SIGUIENTES:")


def tercera_bis(d: dict) -> str:
    return ("TERCERA BIS.-. EL (LA) TRABAJADOR(A) EXTIENDE EL PRESENTE COMO EL FINIQUITO MÁS AMPLIO QUE A DERECHO "
            f"PROCEDA A FAVOR DE {d['empresa']}, UBICADA EN {d['ciudad']}, {d['estado']} CONOCIDA CON NOMBRES "
            f"COMERCIALES COMO {d['sucursal']} Y/O QUIEN REPRESENTE SUS INTERESES, NO RESERVÁNDOSE EL (LA) "
            "TRABAJADOR(A) ACCIÓN NI DERECHO ALGUNO QUE EJERCITAR EN CONTRA DE LAS MENCIONADAS EMPRESAS, NI EN LO "
            "PRESENTE NI EN LO FUTURO EN MATERIA LEGAL ALGUNA, DANDO POR TERMINADA LA RELACIÓN LABORAL QUE LOS UNÍA, Y "
            "POR FINIQUITADA TOTALMENTE DICHA RELACIÓN, ASÍ COMO CUALQUIER PRESTACIÓN DE CARÁCTER LABORAL.")


def asi_mismo(d: dict) -> str:
    return ("Así mismo, continúa manifestando el (la) Trabajador(a) que jamás sufrió ningún tipo de accidente ni "
            f"enfermedad el tiempo que laboró al servicio de {d['empresa']}")


DESTINATARIO = "A QUIEN CORRESPONDA:"
DESPEDIDA = "ATENTAMENTE"
RENUNCIA = [
    "Por medio de la presente le comunico que por convenir así a mis intereses particulares, con fecha del día de hoy "
    "he resuelto de dar terminada voluntariamente la relación laboral y/o contrato individual de trabajo que se me "
    "unía con usted(es) en términos de la fracción I del artículo 53 de la Ley Federal del Trabajo.",
    "Le manifiesto expresamente que durante este tiempo que preste mis servicios, nunca sufrí de riesgo de trabajo "
    "alguno, de igual modo a la fecha no se me adeuda prestación alguna de ningún tipo, y por último y en virtud de "
    "esta renuncia no me reservo acción o derecho de ejercitar de ninguna naturaleza en el futuro, ni en contra suya "
    "ni de su negocio, ni de su prestación legal, ni de ninguna otra persona que hubiese sido mi patrón.",
]
TITULO_CONSTANCIA = "C O N S T A N C I A   L A B O R A L"
CIERRE_CONSTANCIA = ("EXTIENDO LA PRESENTE CARTA LABORAL, SUSCRIBIÉNDOME A SUS ÓRDENES PARA CUALQUIER AMPLIACIÓN A LA "
                     "MISMA.")


def cuerpo_constancia(d: dict) -> str:
    registros = []
    if d["nss"]:
        registros.append(f"NÚMERO DE SEGURIDAD SOCIAL {d['nss']}")
    if d["rfc"]:
        registros.append(f"RFC {d['rfc']}")
    identificacion = ("CON " + " Y ".join(registros) + " ") if registros else ""
    return (f"POR MEDIO DE LA PRESENTE SE HACE CONSTAR QUE EL C. {d['nombre']} LABORÓ PARA LA EMPRESA {d['empresa']} "
            f"{identificacion}A PARTIR DEL DÍA {fecha_texto(d['ingreso'])} "
            f"HASTA EL DÍA {fecha_texto(d['baja'])}.").upper()


# ------------------------------- Maquetado -------------------------------
# Cada documento se convierte en una lista de trazos (texto, líneas, imágenes) medidos en puntos
# sobre una hoja carta, con el origen arriba a la izquierda. Después esa lista se escribe en el PDF.

ANCHO_HOJA, ALTO_HOJA = 612, 792
_LETRAS = {False: "Helvetica", True: "Helvetica-Bold"}


def _tabla_de_anchos() -> dict:
    """Anchos de cada carácter (milésimas del tamaño de letra) para la letra normal y la negrita."""
    tabla = {False: {}, True: {}}
    for codigo in range(32, 256):
        try:
            ch = bytes([codigo]).decode("cp1252")
        except UnicodeDecodeError:
            continue
        for negrita, letra in _LETRAS.items():
            tabla[negrita][ch] = round(pdfmetrics.stringWidth(ch, letra, 1000))
    return tabla


_ANCHOS = _tabla_de_anchos()


def imprimible(s) -> str:
    """Quita lo que la letra del PDF no puede dibujar."""
    return "".join(ch for ch in str(s) if ch in _ANCHOS[False])


def medir(s: str, z: float, negrita: bool = False) -> float:
    tabla = _ANCHOS[bool(negrita)]
    return sum(tabla.get(ch, 556) for ch in s) * z / 1000


class Hoja:
    def __init__(self):
        self.ops: list[dict] = []
        self.fin = 0.0

    def txt(self, s, x, y, z, b=False, c=None):
        s = imprimible(s)
        if s:
            self.ops.append({"t": "txt", "s": s, "x": x, "y": y, "z": z, "b": bool(b), "c": c})

    def centro(self, s, cx, y, z, b=False, c=None):
        s = imprimible(s)
        self.txt(s, cx - medir(s, z, b) / 2, y, z, b, c)

    def der(self, s, xd, y, z, b=False, c=None):
        s = imprimible(s)
        self.txt(s, xd - medir(s, z, b), y, z, b, c)

    def lin(self, x1, y1, x2, y2, w=0.5, c=None):
        self.ops.append({"t": "lin", "x1": x1, "y1": y1, "x2": x2, "y2": y2, "w": w, "c": c})

    def img(self, k, x, y, w, h):
        self.ops.append({"t": "img", "k": k, "x": x, "y": y, "w": w, "h": h})


def _palabras(tramos) -> list[list[tuple[str, bool]]]:
    """Separa en palabras conservando qué fragmentos van en negrita."""
    salida, actual = [], []
    for texto, negrita in tramos:
        for i, parte in enumerate(imprimible(texto).split(" ")):
            if i > 0 and actual:
                salida.append(actual)
                actual = []
            if parte:
                actual.append((parte, bool(negrita)))
    if actual:
        salida.append(actual)
    return salida


def parrafo(h: Hoja, tramos, x, y, ancho, z, alto=None, alin="j", b=False, c=None) -> float:
    """Dibuja un párrafo y devuelve la "y" donde termina. alin: j (justificado), i, c, d."""
    if isinstance(tramos, str):
        tramos = [(tramos, b)]
    alto = alto or z * 1.24
    esp = medir(" ", z)
    pal = [(fr, sum(medir(s, z, n) for s, n in fr)) for fr in _palabras(tramos)]
    i = 0
    while i < len(pal):
        w, j = pal[i][1], i + 1
        while j < len(pal) and w + esp + pal[j][1] <= ancho + 0.01:
            w += esp + pal[j][1]
            j += 1
        n, ultima = j - i, j >= len(pal)
        base = y + (alto - z) / 2 + z * 0.78
        if alin == "j" and not ultima and n > 1:
            hueco = (ancho - (w - esp * (n - 1))) / (n - 1)
            cx = x
            for fr, _ in pal[i:j]:
                for s, neg in fr:
                    h.txt(s, cx, base, z, neg, c)
                    cx += medir(s, z, neg)
                cx += hueco
        else:
            # renglón sin estirar: se junta el texto del mismo estilo en un solo trazo
            cx = x + (ancho - w) / 2 if alin == "c" else x + ancho - w if alin == "d" else x
            buf, bb, bx = "", None, cx
            for k in range(i, j):
                for s, neg in pal[k][0]:
                    if bb != neg:
                        if buf:
                            h.txt(buf, bx, base, z, bb, c)
                        buf, bb, bx = "", neg, cx
                    buf += s
                    cx += medir(s, z, neg)
                if k < j - 1:
                    buf += " "
                    cx += esp
            if buf:
                h.txt(buf, bx, base, z, bb, c)
        y += alto
        i = j
    return y


def _dias(n: float) -> str:
    return ("%.3f" % n).rstrip("0").rstrip(".") if n else ""


def _hoja_convenio(d: dict, n: int, k: float) -> Hoja:
    h, op = Hoja(), d["calc"]["opciones"][n]
    ML = 44
    A = ANCHO_HOJA - 2 * ML
    XD = ML + A
    y = 30

    h.centro(TITULO_CONVENIO, ANCHO_HOJA / 2, y + 12.5 * k, 14.5 * k, True)
    y += 24 * k
    for s in encabezado_convenio(d):
        h.txt(s, ML, y + 10 * k, 10.5 * k, True)
        y += 13.6 * k
    y += 6 * k
    y = parrafo(h, intro_convenio(d), ML, y, A, 9.4 * k, 12 * k)
    y += 6 * k
    h.centro(CLAUSULAS, ANCHO_HOJA / 2, y + 9 * k, 9.4 * k, True)
    y += 17 * k

    z, alto, sep = 8.5 * k, 10.6 * k, 4.5 * k
    y = parrafo(h, primera(d), ML, y, A, z, alto) + sep
    y = parrafo(h, SEGUNDA, ML, y, A, z, alto) + sep
    y = parrafo(h, tercera(d, n), ML, y, A, z, alto) + sep

    # Tabla de percepciones y descuentos
    c1, c2 = ML, ML + A * 0.468
    c3 = c2 + A * 0.21
    c4 = c3 + A * 0.176
    c5 = XD
    zt, hf, g, pad = 6.7 * k, 10.4 * k, 0.4, 3

    def base(yy, tam):
        return yy + hf / 2 + tam * 0.36

    def rejilla(yy):
        h.lin(c1, yy, c5, yy, g)
        h.lin(c1, yy + hf, c5, yy + hf, g)
        for x in (c1, c2, c3, c4, c5):
            h.lin(x, yy, x, yy + hf, g)

    rejilla(y)
    h.txt("PERCEPCIONES:", c1 + pad, base(y, zt), zt, True)
    h.centro("PERIODO:", (c2 + c3) / 2, base(y, zt), zt, True)
    h.centro("Ctd", (c3 + c4) / 2, base(y, zt), zt, True)
    h.centro("Monto", (c4 + c5) / 2, base(y, zt), zt, True)
    y += hf
    for p in d["calc"]["perc"]:
        monto = op["otras"] if p["k"] == "otras" else p["monto"]
        rejilla(y)
        h.txt(p["doc"], c1 + pad, base(y, zt), zt, p["negrita"])
        if p["periodo"]:
            h.centro(p["periodo"], (c2 + c3) / 2, base(y, zt), zt)
        if p["dias"]:
            h.centro(_dias(p["dias"]), (c3 + c4) / 2, base(y, zt), zt)
        if monto:
            h.der(pesos(monto), c5 - pad, base(y, zt), zt)
        y += hf
    h.txt("TOTAL, PERCEPCIONES:", c1 + pad, base(y, zt), zt, True)
    h.der(pesos(op["percepciones"]), c5 - pad, base(y, 7.4 * k), 7.4 * k, True)
    y += hf
    h.lin(c1, y, c5, y, g)
    for x in d["calc"]["desc"]:
        rejilla(y)
        h.txt(x["doc"], c1 + pad, base(y, zt), zt)
        if x["monto"]:
            h.der(pesos(x["monto"]), c5 - pad, base(y, zt), zt)
        y += hf
    h.txt("TOTAL, DEDUCCIONES:", c1 + pad, base(y + 1 * k, zt), zt, True)
    h.der(pesos(op["deducciones"]), c5 - pad, base(y + 1 * k, 8.4 * k), 8.4 * k, True)
    y += hf + 1.5 * k
    h.txt("TOTAL A PAGAR", c1 + pad, base(y + 1 * k, zt), zt, True)
    h.der(pesos(op["total"]), c5 - pad, base(y + 1 * k, 9.8 * k), 9.8 * k, True)
    y += hf + 8 * k

    y = parrafo(h, tercera_bis(d), ML, y, A, z, alto) + sep
    y = parrafo(h, CUARTA, ML, y, A, z, alto) + sep
    y = parrafo(h, asi_mismo(d), ML, y, A, z, alto) + sep
    y = parrafo(h, FUNDAMENTO, ML, y, A, z, alto) + sep
    y = parrafo(h, UNICO, ML, y, A, z, alto)

    # Firmas
    af = A * 0.38
    xi, xd, zf, lf = ML, XD - af, 7 * k, 8.8 * k
    y += 16 * k
    h.centro(EL_TRABAJADOR, xi + af / 2, y + 7 * k, zf, True)
    h.centro(EL_PATRON, xd + af / 2, y + 7 * k, zf, True)
    y += 9 * k + 30 * k
    h.lin(xi, y, xi + af, y, 0.6)
    h.lin(xd, y, xd + af, y, 0.6)
    y += 2.5 * k
    yi = parrafo(h, d["nombre"], xi, y, af, zf, lf, "c", True)
    yd = parrafo(h, d["empresa"], xd, y, af, zf, lf, "c", True)
    yd = parrafo(h, d["representante"], xd, yd + 1 * k, af, zf, lf, "c", True)
    h.fin = max(yi, yd)
    return h


def convenio(d: dict, opciones: list[int]) -> list[Hoja]:
    """Una hoja por cada opción de días. Si no cabe, se reduce la letra poco a poco."""
    limite = ALTO_HOJA - 22
    hojas = []
    for n in opciones:
        k = 1.04
        while True:
            hoja = _hoja_convenio(d, n, k)
            k -= 0.02
            if hoja.fin <= limite or k < 0.699:
                break
        hojas.append(hoja)
    return hojas


def renuncia(d: dict) -> list[Hoja]:
    h = Hoja()
    ML = 52
    A = ANCHO_HOJA - 2 * ML
    XD, CX = ML + A, ANCHO_HOJA / 2
    h.der(lugar_fecha(d), XD, 100, 10.5)
    y = parrafo(h, d["empresa"], ML, 150, A, 12.5, 16, "i")
    y += 34
    h.txt(DESTINATARIO, ML, y + 11, 12.5)
    y += 34
    for i, texto in enumerate(RENUNCIA):
        y = parrafo(h, texto, ML, y, A, 12, 17.5) + (17.5 if i < len(RENUNCIA) - 1 else 0)
    y = max(y + 86, 548)
    h.centro(DESPEDIDA, CX, y + 11, 12.5)
    y += 122
    h.lin(CX - 104, y, CX + 104, y, 0.7)
    parrafo(h, d["nombre"], CX - 150, y + 3, 300, 11, 14, "c", True)
    h.fin = y + 40
    return [h]


def constancia(d: dict) -> list[Hoja]:
    """Constancia laboral en hoja membretada."""
    h, F = Hoja(), d["firmante"]
    ML, XD, CX, gris = 41, 571, ANCHO_HOJA / 2, (127, 127, 127)
    A = XD - ML
    h.img("fondo", 0, 0, ANCHO_HOJA, ALTO_HOJA)
    h.img("logo", 39, 18, 190, 51.8)
    if F["direccion"]:
        h.lin(401, 18, 401, 80, 2.25, (255, 192, 0))
        ya = 19
        for renglon in F["direccion"]:
            ya = parrafo(h, renglon, 411, ya, XD - 411, 9, 11.6, "i", True, gris)
    h.der(lugar_fecha(d), XD, 119, 9.6)
    h.centro(TITULO_CONSTANCIA, CX, 180, 15, True)
    h.txt(DESTINATARIO, ML, 228, 11.3)
    y = parrafo(h, cuerpo_constancia(d), ML, 247, A, 11.2, 15, "i")
    y = parrafo(h, CIERRE_CONSTANCIA, ML, y + 15, A, 11.2, 15)
    y = max(y + 58, 412)
    h.centro(DESPEDIDA, CX, y, 13.4, True)
    yl = y + 106  # renglón de la firma
    firma = F.get("firma")
    if firma and firma.get("ancho") and firma.get("alto"):
        e = min(170 / firma["ancho"], 98 / firma["alto"])
        w, a = firma["ancho"] * e, firma["alto"] * e
        h.img("firma", CX - w / 2, yl + 5 - a, w, a)
    h.lin(CX - 93, yl, CX + 93, yl, 0.7)
    y = parrafo(h, F["nombre"], CX - 180, yl + 3, 360, 11.3, 14.4, "c", True)
    for s in (F["area"], F["telefono"], F["correo"]):
        if s:
            y = parrafo(h, s, CX - 180, y, 360, 11.3, 14.4, "c")
    h.img("sello", CX - 83.5, y + 16, 167, 70)
    h.fin = y + 86
    return [h]


# ---------------------------------- PDF ----------------------------------

def crear_pdf(hojas: list[Hoja], titulo: str, firma_png: bytes | None = None, meta_kw: str = "") -> bytes:
    """Escribe las hojas maquetadas en un PDF tamaño carta y devuelve sus bytes."""
    memoria = io.BytesIO()
    pdf = lienzo_pdf.Canvas(memoria, pagesize=(ANCHO_HOJA, ALTO_HOJA), pageCompression=1)
    pdf.setTitle(titulo)
    pdf.setSubject("Convenios y finiquitos")
    pdf.setAuthor("Recursos Humanos")
    pdf.setCreator("Convenios y finiquitos - Recursos Humanos")
    if meta_kw:
        pdf.setKeywords(meta_kw)
    imagenes = {"fondo": str(RECURSOS / "fondo.jpg"), "logo": str(RECURSOS / "logo.png"),
                "sello": str(RECURSOS / "sello.png")}
    if firma_png:
        imagenes["firma"] = ImageReader(io.BytesIO(firma_png))
    for hoja in hojas:
        for o in hoja.ops:
            r, g, b = (v / 255 for v in (o.get("c") or (0, 0, 0)))
            if o["t"] == "txt":
                pdf.setFont(_LETRAS[o["b"]], o["z"])
                pdf.setFillColorRGB(r, g, b)
                pdf.drawString(o["x"], ALTO_HOJA - o["y"], o["s"])
            elif o["t"] == "lin":
                pdf.setStrokeColorRGB(r, g, b)
                pdf.setLineWidth(o["w"])
                pdf.line(o["x1"], ALTO_HOJA - o["y1"], o["x2"], ALTO_HOJA - o["y2"])
            elif o["t"] == "img" and o["k"] in imagenes:
                pdf.drawImage(imagenes[o["k"]], o["x"], ALTO_HOJA - o["y"] - o["h"], o["w"], o["h"], mask="auto")
        pdf.showPage()
    pdf.save()
    return memoria.getvalue()


def _codificar_meta(documento: str, s: dict, dias) -> str:
    """Empaqueta la captura en ASCII comprimido para reimportar el PDF con fidelidad total."""
    try:
        crudo = json.dumps({"v": 1, "doc": documento, "dias": dias, "s": s}, ensure_ascii=False, default=str)
        comp = zlib.compress(crudo.encode("utf-8"), level=9)
        return "RHFINI1:" + base64.b64encode(comp).decode("ascii")
    except Exception:
        return ""


def _archivo(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|\x00-\x1f]', " ", limpio(s))).strip()[:150]


def nombre_archivo(documento: str, s: dict, dias="0") -> str:
    d = datos(s)
    if documento == "convenio":
        cuales = "TODOS" if str(dias) == "todos" else f"{dias} DÍAS"
        return _archivo(f"FINIQUITO - {d['nombre']} - {d['sucursal']} ({cuales})") + ".pdf"
    prefijo = "RENUNCIA - " if documento == "renuncia" else "CARTA LABORAL - "
    return _archivo(prefijo + d["nombre"]) + ".pdf"


def pdf_documento(documento: str, s: dict, firmante: dict | None = None, dias="0",
                  ahora: datetime | None = None) -> bytes:
    """PDF de "convenio" (dias = 0, 15, 30, 45 o "todos"), "renuncia" o "constancia"."""
    d = datos(s, firmante, ahora)
    titulo = nombre_archivo(documento, s, dias)[:-4]
    meta = _codificar_meta(documento, s, dias)
    if documento == "convenio":
        opciones = DIAS_CONVENIO if str(dias) == "todos" else [int(dias)]
        return crear_pdf(convenio(d, opciones), titulo, meta_kw=meta)
    if documento == "renuncia":
        return crear_pdf(renuncia(d), titulo, meta_kw=meta)
    firma = d["firmante"]["firma"]
    return crear_pdf(constancia(d), titulo, firma["png"] if firma else None, meta_kw=meta)


# --------------------------------- Firma ---------------------------------

def preparar_firma(contenido: bytes) -> dict:
    """Deja la firma como tinta azul sobre fondo transparente y recorta el papel sobrante.

    Devuelve {"png": bytes, "ancho": int, "alto": int}; lanza ValueError si no se distingue la firma.
    """
    try:
        im = ImageOps.exif_transpose(Image.open(io.BytesIO(contenido)))
        im.load()
    except Exception as error:  # archivo dañado o que no es imagen
        raise ValueError("No se pudo abrir la imagen.") from error
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGBA")
        papel_blanco = Image.new("RGBA", im.size, (255, 255, 255, 255))
        papel_blanco.alpha_composite(im)
        im = papel_blanco
    im = im.convert("RGB")
    im.thumbnail((900, 900), Image.LANCZOS)
    luz = im.convert("L")
    cuenta, total = luz.histogram(), luz.width * luz.height

    def percentil(q: float) -> int:
        suma = 0
        for valor, cuantos in enumerate(cuenta):
            suma += cuantos
            if suma >= total * q:
                return valor
        return 255

    papel, tinta = percentil(0.6), percentil(0.002)
    if papel - tinta < 50:
        raise ValueError("No se distingue la firma. Usa una imagen con fondo blanco.")
    corte = papel - (papel - tinta) * 0.22
    rango = corte - tinta
    alfa = luz.point([round(min(1.0, max(0.0, (corte - v) / rango)) * 255) for v in range(256)])
    caja = alfa.point([255 if v > 64 else 0 for v in range(256)]).getbbox()
    if not caja:
        raise ValueError("No se distingue la firma. Usa una imagen con fondo blanco.")
    m = 6
    caja = (max(0, caja[0] - m), max(0, caja[1] - m), min(luz.width, caja[2] + m), min(luz.height, caja[3] + m))
    alfa = alfa.crop(caja)
    firma = Image.new("RGBA", alfa.size, (22, 42, 126, 0))
    firma.putalpha(alfa)
    salida = io.BytesIO()
    firma.save(salida, "PNG", optimize=True)
    return {"png": salida.getvalue(), "ancho": firma.width, "alto": firma.height}


# -------------------------------- Historial ------------------------------
# Mismo acomodo (y 4 renglones por persona: 0, 15, 30 y 45 días) que la hoja "Historico" del Excel.

COLUMNAS_HISTORIAL = [
    "Dias_Negociables", "ID", "NOMBRE", "EMPRESA", "NOMBRE_COMERCIAL", "ESTADO", "CIUDAD", "REPRESENTANTE",
    "F_INGRESO", "FECHAINGRESO", "F_BAJA", "FECHABAJA", "SUELDO", "SUELDO_LETRAS", "Monto_a_Pagar", "Monto_letras",
    "S.D.", "Sueldodiario", "SALARIOS_PENDIENTES", "CTD_SALARIO_PENDIENTES", "DESCFEST",
    "DESCANSO_Y_FESTIVO_LABORADO", "COMISIONES", "AGUINALDO_PROP", "CTD_AGUINALDO_PENDIENTES", "CTD_AGUIN_PENDTES",
    "VACS_PROP", "CTD_VACACIONES_PROPORC", "CTD_VACPROPORC", "PRIMA_VAC_PROPORCIONAL", "CTD_PRIMAVACPROP",
    "VAC_PENDIENTES", "CTDVAC_PENDIENTES", "LIQ_FONDO_AHORRO", "FONDO_RETIRO", "OTRAS_PRESTACIONES_EXCENTAS",
    "DESCUENTO_INFONAVIT", "DESC_FALTANTE", "DESC_CAJA_CHICA", "DESC_PRESTAMO", "DESC_SINIESTRO", "DESC_VIATICOS",
    "DESC_OTROS", "DESC_DIAS_PAGADOS", "TOTAL_DEDUCCIONES", "TOTAL_PERCEPCIONES", "MONTO_TOTAL", "MONTO_TOTAL_PESOS",
]
_COLUMNAS_DE_FECHA = {"F_INGRESO", "F_BAJA"}
_COLUMNAS_DE_TEXTO = {"ID", "NOMBRE", "EMPRESA", "NOMBRE_COMERCIAL", "ESTADO", "CIUDAD", "REPRESENTANTE", "FECHAINGRESO",
                      "FECHABAJA", "SUELDO_LETRAS", "Monto_letras", "Sueldodiario", "MONTO_TOTAL_PESOS"}
_COLUMNAS_DE_CANTIDAD = {"Dias_Negociables", "CTD_SALARIO_PENDIENTES", "DESCFEST", "CTD_AGUINALDO_PENDIENTES",
                         "CTD_AGUIN_PENDTES", "CTD_VACACIONES_PROPORC", "CTD_VACPROPORC", "CTD_PRIMAVACPROP",
                         "CTDVAC_PENDIENTES"}


def filas_historial(s: dict) -> list[list]:
    d = datos(s)
    c = d["calc"]
    P = {p["k"]: p for p in c["perc"]}
    D = {x["k"]: x["monto"] for x in c["desc"]}

    def p(cent):  # centavos -> pesos
        return cent / 100

    filas = []
    for n in DIAS_CONVENIO:
        o = c["opciones"][n]
        filas.append([
            n, d["idNomina"], d["nombre"], d["empresa"], d["sucursal"], d["estado"], d["ciudad"], d["representante"],
            c["ingreso"], fecha_con_dia(c["ingreso"]) if c["ingreso"] else "",
            c["baja"], fecha_con_dia(c["baja"]) if c["baja"] else "",
            p(c["sueldo"]), num_letras(c["sueldo"]).lower(),
            p(c["opciones"][0]["total"]), num_letras(c["opciones"][0]["total"]).lower(),
            c["sueldo"] / 3000, num_letras(c["diario"]).lower(),
            p(P["salarios"]["monto"]), P["salarios"]["dias"], P["descanso"]["dias"], p(P["descanso"]["monto"]), 0,
            p(P["aguinaldo"]["monto"]), P["aguinaldo"]["dias"], round(P["aguinaldo"]["dias"], 2),
            p(P["vacaciones"]["monto"]), P["vacaciones"]["dias"], round(P["vacaciones"]["dias"], 1),
            p(P["prima"]["monto"]), P["prima"]["dias"], p(P["vacPend"]["monto"]), P["vacPend"]["dias"],
            p(P["fondoAhorro"]["monto"]), p(P["fondoRetiro"]["monto"]), p(o["otras"]),
            p(D["infonavit"]), p(D["faltante"]), p(D["cajaChica"]), p(D["prestamo"]), p(D["siniestro"]),
            p(D["viaticos"]), p(D["otros"]), p(D["diasNomina"]),
            p(o["deducciones"]), p(o["percepciones"]), p(o["total"]), num_letras(o["total"]).lower(),
        ])
    return filas


def libro_historial(capturas: list[dict]) -> bytes:
    """Libro de Excel con una hoja "Historico" y cuatro renglones por cada captura."""
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Historico"
    hoja.append(COLUMNAS_HISTORIAL)
    for celda in hoja[1]:
        celda.font = Font(bold=True)
    for s in capturas:
        for fila in filas_historial(s):
            hoja.append(fila)
    for i, columna in enumerate(COLUMNAS_HISTORIAL, start=1):
        hoja.column_dimensions[get_column_letter(i)].width = max(12, min(42, len(columna) + 4))
        if columna in _COLUMNAS_DE_TEXTO or columna in _COLUMNAS_DE_CANTIDAD:
            continue
        formato = "DD/MM/YYYY" if columna in _COLUMNAS_DE_FECHA else "#,##0.00"
        for (celda,) in hoja.iter_rows(min_row=2, min_col=i, max_col=i):
            celda.number_format = formato
    hoja.freeze_panes = "A2"
    memoria = io.BytesIO()
    libro.save(memoria)
    return memoria.getvalue()


# --------------------- Importación de PDF, Excel y Lote ZIP ---------------------

_MESES_MAPA = {sin_acentos(m): i + 1 for i, m in enumerate(MESES)}
_ESTADOS_ORDEN = sorted(ESTADOS, key=len, reverse=True)


def normalizar_captura(parcial: dict | None) -> dict:
    """Garantiza que un diccionario importado tenga todas las llaves esperadas por la pantalla."""
    base = captura_vacia()
    if not parcial:
        return base
    for k in ("empresa", "sucursal", "estado", "ciudad", "representante", "nombre",
              "idNomina", "sueldo", "ingreso", "baja", "hora", "rfc", "nss"):
        val = parcial.get(k)
        if isinstance(val, (date, datetime)):
            base[k] = fecha_corta(val.date() if isinstance(val, datetime) else val)
        else:
            base[k] = "" if val is None else str(val)
    for k in ("incluirHora", "incluirRfc", "incluirNss"):
        if k in parcial:
            base[k] = bool(parcial[k])
    perc_in = parcial.get("perc") or {}
    for p in PERCEPCIONES:
        r = perc_in.get(p["k"]) or {}
        base["perc"][p["k"]] = {
            campo: ("" if r.get(campo) is None else str(r.get(campo)))
            for campo in ("periodo", "dias", "monto")
        }
    desc_in = parcial.get("desc") or {}
    for d in DESCUENTOS:
        val = desc_in.get(d["k"])
        base["desc"][d["k"]] = "" if val is None else str(val)
    return base


def _fecha_larga_a_corta(texto: str) -> str:
    m = re.search(r"(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})", texto, flags=re.IGNORECASE)
    if not m:
        return ""
    dia, mes_txt, anio = int(m.group(1)), _MESES_MAPA.get(sin_acentos(m.group(2))), int(m.group(3))
    if not mes_txt:
        return ""
    try:
        return fecha_corta(date(anio, mes_txt, dia))
    except ValueError:
        return ""


def _emparejar_estado(texto_estado: str) -> str:
    buscado = sin_acentos(texto_estado)
    for e in ESTADOS:
        if sin_acentos(e) == buscado:
            return e
    return may(texto_estado)


def _partir_lugar_fecha(linea: str) -> tuple[str, str, str]:
    """Extrae (ciudad, estado, fecha_baja) de 'CULIACÁN, SINALOA A 26 DE NOVIEMBRE DE 2026.'."""
    m = re.match(r"^(.+?)\s+A\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{4})\.?$", limpio(linea), flags=re.IGNORECASE)
    if not m:
        return "", "", ""
    lugar, f_txt = limpio(m.group(1)), _fecha_larga_a_corta(m.group(2))
    lugar_norm = sin_acentos(lugar)
    for est in _ESTADOS_ORDEN:
        est_norm = sin_acentos(est)
        sufijo = ", " + est_norm
        if lugar_norm.endswith(sufijo):
            ciudad = limpio(lugar[: len(lugar) - len(sufijo)])
            return may(ciudad), est, f_txt
    if "," in lugar:
        ciudad, est_raw = lugar.rsplit(",", 1)
        return may(ciudad), _emparejar_estado(est_raw), f_txt
    return may(lugar), "", f_txt


def leer_pdf_documento(contenido: bytes) -> dict:
    """Lee un PDF de convenio, renuncia o constancia y devuelve sus campos listos para editar."""
    try:
        lector = PdfReader(io.BytesIO(contenido))
    except Exception as error:
        raise ValueError("No se pudo abrir el archivo PDF. Verifica que no esté dañado o protegido.") from error
    if not lector.pages:
        raise ValueError("El PDF no contiene páginas.")

    # 1. Vía exacta: metadatos incrustados por RH Studio
    meta_kw = ""
    try:
        meta_kw = str((lector.metadata or {}).get("/Keywords") or "")
    except Exception:
        meta_kw = ""
    if meta_kw.startswith("RHFINI1:"):
        try:
            raw = zlib.decompress(base64.b64decode(meta_kw[8:].encode("ascii"))).decode("utf-8")
            carga = json.loads(raw)
            dias = carga.get("dias", 0)
            if str(dias) != "todos":
                dias = int(dias) if int(dias) in DIAS_CONVENIO else 0
            return {
                "documento": str(carga.get("doc") or "convenio"),
                "captura": normalizar_captura(carga.get("s")),
                "dias": dias,
                "origen": "metadatos",
            }
        except Exception:
            pass

    # 2. Vía universal: análisis inteligente del texto del PDF (incluye PDFs anteriores)
    texto_p0 = lector.pages[0].extract_text() or ""
    lineas = [limpio(x) for x in texto_p0.splitlines() if limpio(x)]
    plano = " ".join(lineas)
    plano_may = plano.upper()

    s = captura_vacia()
    dias_sel: int | str = 0

    if "CONVENIO ACTO" in plano_may or "PERCEPCIONES:" in plano_may:
        doc_tipo = "convenio"
        m_est = re.search(r"del Estado de (.+?) comparecemos", plano, flags=re.IGNORECASE)
        if not m_est:
            m_est = re.search(r"EN EL ESTADO DE (.+?) P R E S E N T E", plano, flags=re.IGNORECASE)
        if m_est:
            s["estado"] = _emparejar_estado(m_est.group(1))

        m_intro_h = re.search(
            r"En la ciudad de (.+?),\s*siendo las (\d{2}:\d{2}) horas del día (\d{1,2} de [a-záéíóúñ]+ de \d{4}),\s*ante las oficinas",
            plano, flags=re.IGNORECASE,
        )
        m_intro_sin = re.search(
            r"En la ciudad de (.+?),\s*el día (\d{1,2} de [a-záéíóúñ]+ de \d{4}),\s*ante las oficinas",
            plano, flags=re.IGNORECASE,
        )
        lugar_intro = ""
        if m_intro_h:
            lugar_intro = limpio(m_intro_h.group(1))
            s["hora"] = m_intro_h.group(2)
            s["incluirHora"] = True
            s["baja"] = _fecha_larga_a_corta(m_intro_h.group(3))
        elif m_intro_sin:
            lugar_intro = limpio(m_intro_sin.group(1))
            s["hora"] = ""
            s["incluirHora"] = False
            s["baja"] = _fecha_larga_a_corta(m_intro_sin.group(2))

        if lugar_intro:
            if s["estado"] and sin_acentos(lugar_intro).endswith(", " + sin_acentos(s["estado"])):
                s["ciudad"] = may(lugar_intro[: len(lugar_intro) - len(s["estado"]) - 2])
            elif "," in lugar_intro:
                ciu, est_r = lugar_intro.rsplit(",", 1)
                s["ciudad"] = may(ciu)
                if not s["estado"]:
                    s["estado"] = _emparejar_estado(est_r)
            else:
                s["ciudad"] = may(lugar_intro)

        m_partes = re.search(
            r"comparecemos los suscritos C\.\s*(.+?)\s+como\s+[“\"']?TRABAJADOR\(A\)[”\"']?\s+y por la otra parte la empresa\s+(.+?)\s+como\s+[“\"']?PATR[ÓO]N[”\"']?,\s*representada en este acto por el \(la\) C\.\s*(.+?)\s+quienes de com[úu]n acuerdo",
            plano, flags=re.IGNORECASE,
        )
        if m_partes:
            s["nombre"] = may(m_partes.group(1))
            s["empresa"] = may(m_partes.group(2))
            s["representante"] = may(m_partes.group(3))

        m_prim = re.search(
            r"inici[óo] a partir del d[íi]a (\d{1,2} de [a-záéíóúñ]+ de \d{4}) hasta el d[íi]a (\d{1,2} de [a-záéíóúñ]+ de \d{4}).*?salario diario de \$([\d,]+\.\d{2})",
            plano, flags=re.IGNORECASE,
        )
        diario_cent = 0
        if m_prim:
            s["ingreso"] = _fecha_larga_a_corta(m_prim.group(1))
            if not s["baja"]:
                s["baja"] = _fecha_larga_a_corta(m_prim.group(2))
            diario_cent = centavos(m_prim.group(3))
            if diario_cent:
                s["sueldo"] = numero(diario_cent * 30)

        m_suc = re.search(
            r"CONOCIDA CON NOMBRES COMERCIALES COMO (.+?) Y/O QUIEN REPRESENTE SUS INTERESES",
            plano, flags=re.IGNORECASE,
        )
        if m_suc:
            s["sucursal"] = may(m_suc.group(1))

        # Tabla de percepciones
        etiquetas_perc = [p["doc"] for p in PERCEPCIONES] + ["TOTAL, PERCEPCIONES:"]
        pos_perc = {}
        cursor = 0
        for etq in etiquetas_perc:
            for idx in range(cursor, len(lineas)):
                if sin_acentos(lineas[idx]) == sin_acentos(etq):
                    pos_perc[etq] = idx
                    cursor = idx + 1
                    break

        for i, p in enumerate(PERCEPCIONES):
            ini = pos_perc.get(p["doc"])
            fin = pos_perc.get(etiquetas_perc[i + 1])
            if ini is None or fin is None or fin <= ini:
                continue
            celdas = lineas[ini + 1 : fin]
            montos_tok = [t for t in celdas if "$" in t]
            otros_tok = [t for t in celdas if "$" not in t]
            if montos_tok:
                c_val = centavos(montos_tok[-1])
                s["perc"][p["k"]]["monto"] = numero(c_val)
            elif p.get("obligatorio"):
                s["perc"][p["k"]]["monto"] = "0.00"
            if p.get("periodo") and p.get("dias"):
                if len(otros_tok) >= 2:
                    s["perc"][p["k"]]["periodo"] = otros_tok[0]
                    s["perc"][p["k"]]["dias"] = otros_tok[1]
                elif len(otros_tok) == 1:
                    if re.fullmatch(r"20\d{2}", otros_tok[0]):
                        s["perc"][p["k"]]["periodo"] = otros_tok[0]
                    else:
                        s["perc"][p["k"]]["dias"] = otros_tok[0]
            elif p.get("dias") and otros_tok:
                s["perc"][p["k"]]["dias"] = otros_tok[0]

        # Tabla de descuentos
        etiquetas_desc = [d["doc"] for d in DESCUENTOS] + ["TOTAL, DEDUCCIONES:"]
        pos_desc = {}
        for etq in etiquetas_desc:
            for idx in range(cursor, len(lineas)):
                if sin_acentos(lineas[idx]) == sin_acentos(etq):
                    pos_desc[etq] = idx
                    cursor = idx + 1
                    break

        for i, d_item in enumerate(DESCUENTOS):
            ini = pos_desc.get(d_item["doc"])
            fin = pos_desc.get(etiquetas_desc[i + 1])
            if ini is None or fin is None or fin <= ini:
                continue
            montos_tok = [t for t in lineas[ini + 1 : fin] if "$" in t]
            if montos_tok:
                s["desc"][d_item["k"]] = numero(centavos(montos_tok[-1]))

        # Detección inteligente de días de convenio (0, 15, 30, 45 o 4 hojas)
        if len(lector.pages) == 4:
            dias_sel = "todos"
        elif "PAGO REALIZADO EN EFECTIVO" in plano_may and diario_cent > 0:
            otras_c = centavos(s["perc"]["otras"]["monto"])
            sueldo_c = diario_cent * 30
            for cand in (15, 30, 45):
                if otras_c == dividir(sueldo_c * cand, 30):
                    dias_sel = cand
                    s["perc"]["otras"]["monto"] = ""
                    break

        return {"documento": doc_tipo, "captura": s, "dias": dias_sel, "origen": "texto"}

    if "C O N S T A N C I A" in plano_may or "SE HACE CONSTAR QUE" in plano_may:
        doc_tipo = "constancia"
        for lin in lineas[:10]:
            if re.search(r"\s+A\s+\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{4}\.?$", lin, flags=re.IGNORECASE):
                s["ciudad"], s["estado"], s["baja"] = _partir_lugar_fecha(lin)
                break
        m_cuerpo = re.search(
            r"SE HACE CONSTAR QUE EL C\.\s*(.+?)\s+LABOR[ÓO] PARA LA EMPRESA\s+(.+?)\s+A PARTIR DEL D[ÍI]A\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{4})\s+HASTA EL D[ÍI]A\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{4})",
            plano, flags=re.IGNORECASE,
        )
        if m_cuerpo:
            s["nombre"] = may(m_cuerpo.group(1))
            emp_bloque = limpio(m_cuerpo.group(2))
            s["ingreso"] = _fecha_larga_a_corta(m_cuerpo.group(3))
            if not s["baja"]:
                s["baja"] = _fecha_larga_a_corta(m_cuerpo.group(4))

            m_ambos = re.match(
                r"^(.+?)\s+CON N[ÚU]MERO DE SEGURIDAD SOCIAL\s+(\d{11})\s+Y RFC\s+([A-ZÑ&0-9]{12,13})$",
                emp_bloque, flags=re.IGNORECASE,
            )
            m_solo_nss = re.match(
                r"^(.+?)\s+CON N[ÚU]MERO DE SEGURIDAD SOCIAL\s+(\d{11})$",
                emp_bloque, flags=re.IGNORECASE,
            )
            m_solo_rfc = re.match(
                r"^(.+?)\s+CON RFC\s+([A-ZÑ&0-9]{12,13})$",
                emp_bloque, flags=re.IGNORECASE,
            )
            if m_ambos:
                s["empresa"] = may(m_ambos.group(1))
                s["nss"], s["rfc"] = m_ambos.group(2), may(m_ambos.group(3))
                s["incluirNss"], s["incluirRfc"] = True, True
            elif m_solo_nss:
                s["empresa"] = may(m_solo_nss.group(1))
                s["nss"], s["incluirNss"], s["incluirRfc"] = m_solo_nss.group(2), True, False
            elif m_solo_rfc:
                s["empresa"] = may(m_solo_rfc.group(1))
                s["rfc"], s["incluirRfc"], s["incluirNss"] = may(m_solo_rfc.group(2)), True, False
            else:
                s["empresa"] = may(emp_bloque)
                s["incluirRfc"], s["incluirNss"] = False, False
        return {"documento": doc_tipo, "captura": s, "dias": 0, "origen": "texto"}

    if "INTERESES PARTICULARES" in plano_may or "DAR TERMINADA VOLUNTARIAMENTE" in plano_may:
        doc_tipo = "renuncia"
        if lineas:
            s["ciudad"], s["estado"], s["baja"] = _partir_lugar_fecha(lineas[0])
        if "A QUIEN CORRESPONDA:" in lineas:
            idx_dest = lineas.index("A QUIEN CORRESPONDA:")
            if idx_dest > 1:
                s["empresa"] = may(" ".join(lineas[1:idx_dest]))
        if "ATENTAMENTE" in lineas:
            idx_att = lineas.index("ATENTAMENTE")
            if idx_att + 1 < len(lineas):
                s["nombre"] = may(" ".join(lineas[idx_att + 1 :]))
        return {"documento": doc_tipo, "captura": s, "dias": 0, "origen": "texto"}

    raise ValueError("Este PDF no coincide con el formato de Convenio, Carta de renuncia ni Constancia laboral.")


def leer_excel_historial(contenido: bytes) -> list[dict]:
    """Reconstruye las capturas desde un archivo Excel 'Historico convenios.xlsx'."""
    try:
        wb = load_workbook(io.BytesIO(contenido), data_only=True)
    except Exception as error:
        raise ValueError("No se pudo abrir el archivo Excel.") from error
    ws = wb["Historico"] if "Historico" in wb.sheetnames else wb.active
    filas = list(ws.iter_rows(values_only=True))
    if len(filas) < 2:
        return []
    cabecera = [limpio(c) for c in filas[0]]
    idx = {col: i for i, col in enumerate(cabecera) if col}
    if "NOMBRE" not in idx or "EMPRESA" not in idx:
        raise ValueError("El archivo Excel no tiene las columnas de la hoja Historico.")

    def val(fila, col):
        pos = idx.get(col)
        return fila[pos] if pos is not None and pos < len(fila) else None

    def fmt_m(v, oblig=False):
        c = centavos(v)
        if c:
            return numero(c)
        return "0.00" if oblig else ""

    def fmt_d(v):
        try:
            n = a_numero(v)
            return _dias(n) if n else ""
        except ValueError:
            return ""

    capturas = []
    for fila in filas[1:]:
        dias_neg = val(fila, "Dias_Negociables")
        if dias_neg not in (0, "0", None, "") and str(dias_neg).strip() != "0":
            continue
        s = captura_vacia()
        s["idNomina"] = limpio(val(fila, "ID"))
        s["nombre"] = may(val(fila, "NOMBRE"))
        s["empresa"] = may(val(fila, "EMPRESA"))
        s["sucursal"] = may(val(fila, "NOMBRE_COMERCIAL"))
        s["estado"] = _emparejar_estado(limpio(val(fila, "ESTADO"))) if limpio(val(fila, "ESTADO")) else ""
        s["ciudad"] = may(val(fila, "CIUDAD"))
        s["representante"] = may(val(fila, "REPRESENTANTE"))
        f_ing, f_baj = a_fecha(val(fila, "F_INGRESO")), a_fecha(val(fila, "F_BAJA"))
        s["ingreso"] = fecha_corta(f_ing) if f_ing else ""
        s["baja"] = fecha_corta(f_baj) if f_baj else ""
        s["sueldo"] = fmt_m(val(fila, "SUELDO"))

        s["perc"]["salarios"] = {"periodo": "", "dias": fmt_d(val(fila, "CTD_SALARIO_PENDIENTES")), "monto": fmt_m(val(fila, "SALARIOS_PENDIENTES"))}
        s["perc"]["descanso"] = {"periodo": "", "dias": fmt_d(val(fila, "DESCFEST")), "monto": fmt_m(val(fila, "DESCANSO_Y_FESTIVO_LABORADO"))}
        s["perc"]["aguinaldo"] = {"periodo": "", "dias": fmt_d(val(fila, "CTD_AGUINALDO_PENDIENTES")), "monto": fmt_m(val(fila, "AGUINALDO_PROP"), True)}
        s["perc"]["vacaciones"] = {"periodo": "", "dias": fmt_d(val(fila, "CTD_VACACIONES_PROPORC")), "monto": fmt_m(val(fila, "VACS_PROP"), True)}
        s["perc"]["prima"] = {"periodo": "", "dias": fmt_d(val(fila, "CTD_PRIMAVACPROP")), "monto": fmt_m(val(fila, "PRIMA_VAC_PROPORCIONAL"), True)}
        s["perc"]["vacPend"] = {"periodo": "", "dias": fmt_d(val(fila, "CTDVAC_PENDIENTES")), "monto": fmt_m(val(fila, "VAC_PENDIENTES"))}
        s["perc"]["fondoAhorro"] = {"periodo": "", "dias": "", "monto": fmt_m(val(fila, "LIQ_FONDO_AHORRO"))}
        s["perc"]["fondoRetiro"] = {"periodo": "", "dias": "", "monto": fmt_m(val(fila, "FONDO_RETIRO"))}
        s["perc"]["otras"] = {"periodo": "", "dias": "", "monto": fmt_m(val(fila, "OTRAS_PRESTACIONES_EXCENTAS"))}

        mapa_desc = {
            "infonavit": "DESCUENTO_INFONAVIT", "faltante": "DESC_FALTANTE", "cajaChica": "DESC_CAJA_CHICA",
            "prestamo": "DESC_PRESTAMO", "siniestro": "DESC_SINIESTRO", "viaticos": "DESC_VIATICOS",
            "otros": "DESC_OTROS", "diasNomina": "DESC_DIAS_PAGADOS",
        }
        for k_desc, col_xl in mapa_desc.items():
            s["desc"][k_desc] = fmt_m(val(fila, col_xl))

        if s["nombre"] or s["empresa"]:
            capturas.append(s)
    return capturas


def paquete_lote_zip(elementos: list[tuple[dict, int | str]], firmante: dict | None = None) -> bytes:
    """Empaqueta los PDFs habilitados de varias hojas junto con el Excel consolidado en un archivo ZIP."""
    memoria = io.BytesIO()
    capturas_convenio = []
    usados: set[str] = set()

    def unico(nombre: str) -> str:
        base, ext = (nombre[:-4], ".pdf") if nombre.lower().endswith(".pdf") else (nombre, "")
        cand = nombre
        n = 2
        while cand in usados:
            cand = f"{base} ({n}){ext}"
            n += 1
        usados.add(cand)
        return cand

    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as zf:
        for s, dias in elementos:
            if not validar(s, firmante, "convenio"):
                zf.writestr(unico(nombre_archivo("convenio", s, dias)), pdf_documento("convenio", s, firmante, dias))
                capturas_convenio.append(s)
            if not validar(s, firmante, "renuncia"):
                zf.writestr(unico(nombre_archivo("renuncia", s, dias)), pdf_documento("renuncia", s, firmante, dias))
            if firmante and not validar(s, firmante, "constancia"):
                zf.writestr(unico(nombre_archivo("constancia", s, dias)), pdf_documento("constancia", s, firmante, dias))
        if capturas_convenio:
            zf.writestr(f"Historico convenios {ahora_local():%Y-%m-%d}.xlsx", libro_historial(capturas_convenio))
    return memoria.getvalue()
