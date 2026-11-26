"""Convenios y finiquitos · app de Streamlit para Recursos Humanos.

Para probarla en tu equipo:   streamlit run app.py

Incluye:
- Mesa de trabajo multi-hoja (tipo pestañas de Excel) para capturar varias bajas en paralelo.
- Importador inteligente de PDF (Convenio, Renuncia o Constancia) para editar 1 o 2 datos sin recapturar todo.
- Centro de Historial persistente con buscador, exportación oficial a Excel y restauración desde Excel.
"""
from __future__ import annotations

import base64
import copy
import hmac
import json
import time
import uuid
from pathlib import Path

import streamlit as st

import documentos as D
import ui

st.set_page_config(
    page_title="Convenios y finiquitos · RH Studio",
    page_icon=":material/description:",
    layout="wide",
    initial_sidebar_state=365,
)
ui.apply_styles()

ARCHIVO_HISTORIAL_SERVIDOR = Path(__file__).resolve().parent / ".historial_rh_local.json"

TITULOS = {"convenio": "Convenio de finiquito", "renuncia": "Carta de renuncia", "constancia": "Constancia laboral"}
BOTONES = {"convenio": "Ver y descargar el convenio", "renuncia": "Ver y descargar la renuncia",
           "constancia": "Ver y descargar la constancia"}
OPCIONES_DIAS = [0, 15, 30, 45, "todos"]
# Los rótulos de las opciones deben ser fijos: Streamlit identifica la opción elegida por su texto.
ROTULOS_DIAS = {0: "Sin días", 15: "15 días", 30: "30 días", 45: "45 días", "todos": "Las cuatro opciones"}
DETALLE_DIAS = {0: "sin días de convenio", 15: "con 15 días", 30: "con 30 días", 45: "con 45 días",
                "todos": "las cuatro opciones"}

TEXTOS = ("representante", "sucursal", "ciudad", "nombre", "idNomina", "ingreso", "baja", "sueldo", "hora", "rfc", "nss")
OPCIONALES = ("incluirHora", "incluirRfc", "incluirNss")


def md(texto: str) -> str:
    """Escapa el signo de pesos para que Streamlit no lo tome como fórmula al mostrar texto."""
    return str(texto).replace("$", "\\$")


# ------------------------------- Acceso ----------------------------------

def secreto(nombre: str, defecto=""):
    """Lee un valor de .streamlit/secrets.toml (o de los secretos de la nube) sin fallar si no existe."""
    try:
        return st.secrets.get(nombre, defecto)
    except Exception:  # no hay archivo de secretos
        return defecto


def pedir_clave() -> None:
    """Si en los secretos hay una `clave`, la pide antes de mostrar la página."""
    clave = str(secreto("clave") or "")
    if not clave or st.session_state.get("_autorizado"):
        return
    ui.auth_header()
    _, centro, _ = st.columns([1.2, 1.5, 1.2])
    with centro, st.container(key="sec_empresa"):
        with st.form("acceso", border=False):
            intento = st.text_input("Clave de acceso", type="password", placeholder="Ingresa la clave de RH")
            if st.form_submit_button("Entrar al estudio", type="primary", icon=":material/lock_open:", width="stretch"):
                if hmac.compare_digest(intento.encode(), clave.encode()):
                    st.session_state["_autorizado"] = True
                    st.rerun()
                time.sleep(1.5)  # frena a quien intente adivinarla probando muchas
                st.error("La clave no coincide.")
    st.stop()


# --------------------- Persistencia local y Hojas ------------------------

def nueva_estructura_hoja(captura: dict | None = None, dias=0) -> dict:
    dias_val = dias if dias in OPCIONES_DIAS else 0
    return {
        "id": uuid.uuid4().hex,
        "captura": D.normalizar_captura(captura),
        "dias": dias_val,
    }


def firma_de_archivo(ruta) -> dict | None:
    """Firma por defecto: ruta a una imagen indicada en los secretos (opcional)."""
    if not ruta:
        return None
    archivo = Path(__file__).resolve().parent / str(ruta)
    try:
        return D.preparar_firma(archivo.read_bytes())
    except (OSError, ValueError):
        return None


def leer_historial_servidor() -> list[dict]:
    """Lee la bitácora compartida en el contenedor activo (local o Streamlit Cloud)."""
    if not ARCHIVO_HISTORIAL_SERVIDOR.exists():
        return []
    try:
        crudo = json.loads(ARCHIVO_HISTORIAL_SERVIDOR.read_text(encoding="utf-8"))
        if not isinstance(crudo, list):
            return []
        return [
            {"id": str(r.get("id") or uuid.uuid4().hex), "captura": D.normalizar_captura(r.get("captura"))}
            for r in crudo
            if isinstance(r, dict) and r.get("captura")
        ]
    except Exception:
        return []


def sincronizar_historial_servidor(reemplazar_todo: bool = False) -> None:
    """Fusiona el historial de esta sesión con el del servidor sin pisar registros de otros usuarios."""
    ss = st.session_state
    try:
        if reemplazar_todo:
            final = [
                {"id": r["id"], "captura": D.normalizar_captura(r["captura"])}
                for r in ss.get("historial", [])
            ]
        else:
            remoto = leer_historial_servidor()
            locales = ss.get("historial", [])
            por_id = {r["id"]: {"id": r["id"], "captura": D.normalizar_captura(r["captura"])} for r in locales}
            combinado = [por_id[r["id"]] for r in locales]
            for r in remoto:
                if r["id"] not in por_id:
                    combinado.append(r)
            final = combinado
        ss["historial"] = final
        ARCHIVO_HISTORIAL_SERVIDOR.write_text(json.dumps(final, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def guardar_estado_disco() -> None:
    """Sincroniza el historial compartido sin mezclar los borradores en curso entre usuarios."""
    sincronizar_historial_servidor(reemplazar_todo=False)


def iniciar_estado() -> None:
    ss = st.session_state
    if "_inicializado" in ss:
        # Refresca en cada corrida posibles convenios agregados por otro miembro del equipo
        remoto = leer_historial_servidor()
        if remoto:
            ids_locales = {r["id"] for r in ss.get("historial", [])}
            for r in remoto:
                if r["id"] not in ids_locales:
                    ss["historial"].append(r)
        return

    f_sec = dict(secreto("firmante", {}) or {})
    firmante_defecto = {
        "nombre": D.may(f_sec.get("nombre")),
        "area": D.may(f_sec.get("area") or "RECURSOS HUMANOS"),
        "telefono": str(f_sec.get("telefono") or ""),
        "correo": str(f_sec.get("correo") or ""),
        "direccion": str(f_sec.get("direccion") or ""),
    }

    primera = nueva_estructura_hoja()
    ss["hojas"] = [primera]
    ss["hoja_activa_id"] = primera["id"]
    ss["historial"] = leer_historial_servidor()
    ss["_firmante_txt"] = firmante_defecto

    ss["firma"] = ss["_firma_base"] = firma_de_archivo(f_sec.get("firma"))
    ss["_firma_huella"] = None
    ss["_firma_error"] = ""
    ss["vista_activa"] = "captura"
    ss["_hidratar_pantalla"] = True
    ss["_toast_pendiente"] = ""
    ss["_inicializado"] = True


def hoja_activa() -> dict:
    ss = st.session_state
    for h in ss["hojas"]:
        if h["id"] == ss["hoja_activa_id"]:
            return h
    ss["hoja_activa_id"] = ss["hojas"][0]["id"]
    return ss["hojas"][0]


def indice_hoja_activa() -> int:
    ss = st.session_state
    for idx, h in enumerate(ss["hojas"], start=1):
        if h["id"] == ss["hoja_activa_id"]:
            return idx
    return 1


def hidratar_widgets_desde_estado() -> None:
    """Carga en las claves de los widgets los datos de la hoja activa antes de dibujarlos."""
    ss = st.session_state
    h = hoja_activa()
    s = D.normalizar_captura(h["captura"])
    ss["empresa"] = s["empresa"] or None
    ss["estado"] = s["estado"] if s["estado"] in D.ESTADOS else None
    for k in TEXTOS:
        ss[k] = s.get(k) or ""
    for k in OPCIONALES:
        ss[k] = bool(s.get(k, True))
    for p in D.PERCEPCIONES:
        for campo in ("periodo", "dias", "monto"):
            ss[f"perc.{p['k']}.{campo}"] = s["perc"][p["k"]].get(campo, "")
    for d in D.DESCUENTOS:
        ss[f"desc.{d['k']}"] = s["desc"].get(d["k"], "")
    ss["dias"] = h["dias"] if h.get("dias") in OPCIONES_DIAS else 0

    ft = ss["_firmante_txt"]
    ss["f_nombre"] = ft.get("nombre", "")
    ss["f_area"] = ft.get("area", "RECURSOS HUMANOS")
    ss["f_telefono"] = ft.get("telefono", "")
    ss["f_correo"] = ft.get("correo", "")
    ss["f_direccion"] = ft.get("direccion", "")


def leer_captura_widgets() -> dict:
    """Lee lo escrito actualmente en los widgets de captura."""
    ss = st.session_state
    s = D.captura_vacia()
    for k in ("empresa", "estado") + TEXTOS:
        s[k] = ss.get(k) or ""
    for k in OPCIONALES:
        s[k] = bool(ss.get(k, True))
    for p in D.PERCEPCIONES:
        s["perc"][p["k"]] = {campo: ss.get(f"perc.{p['k']}.{campo}", "") for campo in ("periodo", "dias", "monto")}
    for d in D.DESCUENTOS:
        s["desc"][d["k"]] = ss.get(f"desc.{d['k']}", "")
    return s


def respaldar_widgets_en_hoja_activa() -> None:
    """Si los widgets de captura están montados, vuelca sus valores en la hoja activa."""
    ss = st.session_state
    if ss.get("vista_activa") == "captura" and "nombre" in ss and not ss.get("_hidratar_pantalla"):
        h = hoja_activa()
        h["captura"] = leer_captura_widgets()
        if ss.get("dias") in OPCIONES_DIAS:
            h["dias"] = ss["dias"]
        if "f_nombre" in ss:
            ss["_firmante_txt"] = {
                "nombre": ss.get("f_nombre", ""),
                "area": ss.get("f_area", "RECURSOS HUMANOS"),
                "telefono": ss.get("f_telefono", ""),
                "correo": ss.get("f_correo", ""),
                "direccion": ss.get("f_direccion", ""),
            }


def leer_firmante() -> dict:
    ss = st.session_state
    ft = ss["_firmante_txt"]
    return {
        "nombre": ft.get("nombre", ""),
        "area": ft.get("area", "RECURSOS HUMANOS"),
        "telefono": ft.get("telefono", ""),
        "correo": ft.get("correo", ""),
        "direccion": ft.get("direccion", ""),
        "firma": ss.get("firma"),
    }


def hay_captura(s: dict) -> bool:
    """True si ya se escribió algo en la hoja."""
    vacia = D.captura_vacia()
    sueltos = any(D.limpio(s.get(k)) for k in vacia if k not in ("perc", "desc") + OPCIONALES)
    return sueltos or s.get("perc") != vacia["perc"] or s.get("desc") != vacia["desc"]


def resumen_antiguedad(calc: dict) -> str:
    ing, baj = calc.get("ingreso"), calc.get("baja")
    if ing and baj and baj >= ing:
        dias = (baj - ing).days + 1
        anios, resto = divmod(dias, 365)
        meses = resto // 30
        if anios >= 1:
            return f"Antigüedad: {anios}a {meses}m ({dias:,} días)"
        return f"Periodo: {dias:,} días laborados"
    if calc.get("sueldo"):
        return f"Mensual: {D.pesos(calc['sueldo'])}"
    return "Sueldo mensual ÷ 30 días"


# ----------------------- Acciones de Hojas y Vistas ----------------------

def cambiar_vista(nueva_vista: str) -> None:
    respaldar_widgets_en_hoja_activa()
    st.session_state["vista_activa"] = nueva_vista
    st.session_state["_hidratar_pantalla"] = True
    guardar_estado_disco()


def seleccionar_hoja(id_destino: str) -> None:
    respaldar_widgets_en_hoja_activa()
    st.session_state["hoja_activa_id"] = id_destino
    st.session_state["vista_activa"] = "captura"
    st.session_state["_hidratar_pantalla"] = True
    guardar_estado_disco()


def crear_hoja_vacia() -> None:
    respaldar_widgets_en_hoja_activa()
    actual = hoja_activa()["captura"]
    base = D.captura_vacia()
    # Conserva las preferencias de interruptores opcionales para capturar más rápido
    for k in OPCIONALES:
        base[k] = bool(actual.get(k, True))
    nueva = nueva_estructura_hoja(base, dias=0)
    st.session_state["hojas"].append(nueva)
    st.session_state["hoja_activa_id"] = nueva["id"]
    st.session_state["vista_activa"] = "captura"
    st.session_state["_hidratar_pantalla"] = True
    guardar_estado_disco()


def crear_hoja_misma_empresa() -> None:
    """Crea una nueva hoja copiando únicamente empresa, representante, sucursal, estado y ciudad."""
    respaldar_widgets_en_hoja_activa()
    actual = hoja_activa()["captura"]
    base = D.captura_vacia()
    for k in ("empresa", "representante", "sucursal", "estado", "ciudad"):
        base[k] = actual.get(k, "")
    for k in OPCIONALES:
        base[k] = bool(actual.get(k, True))
    nueva = nueva_estructura_hoja(base, dias=0)
    st.session_state["hojas"].append(nueva)
    st.session_state["hoja_activa_id"] = nueva["id"]
    st.session_state["vista_activa"] = "captura"
    st.session_state["_hidratar_pantalla"] = True
    st.session_state["_toast_pendiente"] = "Nueva hoja lista con la misma empresa y sucursal."
    guardar_estado_disco()


def duplicar_hoja_actual() -> None:
    """Clona todos los datos de la hoja actual en una pestaña nueva para cambiar solo 1 o 2 campos."""
    respaldar_widgets_en_hoja_activa()
    actual = hoja_activa()
    clon = nueva_estructura_hoja(copy.deepcopy(actual["captura"]), dias=actual.get("dias", 0))
    st.session_state["hojas"].append(clon)
    st.session_state["hoja_activa_id"] = clon["id"]
    st.session_state["vista_activa"] = "captura"
    st.session_state["_hidratar_pantalla"] = True
    st.session_state["_toast_pendiente"] = "Hoja duplicada: edita los datos necesarios."
    guardar_estado_disco()


def cerrar_hoja_actual() -> None:
    ss = st.session_state
    if len(ss["hojas"]) <= 1:
        limpiar_hoja_actual()
        return
    id_borrar = ss["hoja_activa_id"]
    idx = next((i for i, h in enumerate(ss["hojas"]) if h["id"] == id_borrar), 0)
    ss["hojas"] = [h for h in ss["hojas"] if h["id"] != id_borrar]
    nuevo_idx = min(idx, len(ss["hojas"]) - 1)
    ss["hoja_activa_id"] = ss["hojas"][nuevo_idx]["id"]
    ss["_hidratar_pantalla"] = True
    guardar_estado_disco()


def limpiar_hoja_actual() -> None:
    h = hoja_activa()
    base = D.captura_vacia()
    for k in OPCIONALES:
        base[k] = bool(h["captura"].get(k, True))
    h["id"] = uuid.uuid4().hex
    h["captura"] = base
    h["dias"] = 0
    st.session_state["hoja_activa_id"] = h["id"]
    st.session_state["_hidratar_pantalla"] = True
    guardar_estado_disco()


def abrir_captura_en_nueva_hoja(captura: dict, dias=0, mensaje: str = "") -> None:
    respaldar_widgets_en_hoja_activa()
    # Si la única hoja abierta está completamente vacía, la aprovecha en vez de dejar una hoja en blanco
    ss = st.session_state
    if len(ss["hojas"]) == 1 and not hay_captura(ss["hojas"][0]["captura"]):
        ss["hojas"][0]["captura"] = D.normalizar_captura(captura)
        ss["hojas"][0]["dias"] = dias if dias in OPCIONES_DIAS else 0
        ss["hoja_activa_id"] = ss["hojas"][0]["id"]
    else:
        nueva = nueva_estructura_hoja(captura, dias=dias)
        ss["hojas"].append(nueva)
        ss["hoja_activa_id"] = nueva["id"]
    ss["vista_activa"] = "captura"
    ss["_hidratar_pantalla"] = True
    if mensaje:
        ss["_toast_pendiente"] = mensaje
    guardar_estado_disco()


# ----------------------- Acciones de Campos e Historial ------------------

def a_mayusculas(clave: str) -> None:
    st.session_state[clave] = D.may(st.session_state[clave])


def completar_empresa() -> None:
    """Si lo escrito coincide con una sola razón social del catálogo, elige esa; si no, lo deja en mayúsculas."""
    valor = D.may(st.session_state.get("empresa"))
    if not valor:
        return
    if valor not in D.EMPRESAS:
        buscado = D.sin_acentos(valor)
        empiezan = [e for e in D.EMPRESAS if D.sin_acentos(e).startswith(buscado)]
        contienen = [e for e in D.EMPRESAS if buscado in D.sin_acentos(e)]
        if len(empiezan) == 1:
            valor = empiezan[0]
        elif len(contienen) == 1:
            valor = contienen[0]
    st.session_state["empresa"] = valor


def dar_formato(clave: str) -> None:
    """Al salir de un monto lo deja con comas y centavos: 7274.01 -> 7,274.01."""
    try:
        c = D.a_centavos(st.session_state[clave])
    except ValueError:
        return
    if c is not None:
        st.session_state[clave] = D.numero(c)


def formato_fecha(clave: str) -> None:
    """Al salir de una fecha la deja como 20/02/2026 (se puede escribir 20-2-26 o 20022026)."""
    f = D.a_fecha(st.session_state[clave])
    if f:
        st.session_state[clave] = D.fecha_corta(f)


def anotar_en_historial(captura: dict, identificador: str) -> None:
    """Guarda (o actualiza) la captura en el historial persistente."""
    historial = st.session_state["historial"]
    limpia = D.normalizar_captura(captura)
    historial[:] = [{"id": identificador, "captura": limpia}] + [r for r in historial if r["id"] != identificador]
    guardar_estado_disco()


def guardar_hoja_actual_en_historial() -> None:
    respaldar_widgets_en_hoja_activa()
    h = hoja_activa()
    if not hay_captura(h["captura"]):
        st.session_state["_toast_pendiente"] = "Escribe al menos el nombre o empresa antes de guardar en el historial."
        return
    anotar_en_historial(h["captura"], h["id"])
    st.session_state["_toast_pendiente"] = "Hoja guardada en el historial."


# ------------------------------ Controles --------------------------------

def texto(etiqueta: str, clave: str, lugar=st, mayusculas: bool = True, **extra) -> None:
    if mayusculas:
        extra.update(on_change=a_mayusculas, args=(clave,))
    lugar.text_input(etiqueta, key=clave, **extra)


def avisar(lugar, avisos: dict, clave: str) -> bool:
    if clave not in avisos:
        return False
    lugar.caption(md(f":red[{avisos[clave]}]"))
    return True


def fecha(etiqueta: str, clave: str, lugar, calc: dict, avisos: dict, ayuda: str = "") -> None:
    lugar.text_input(etiqueta, key=clave, placeholder="DD/MM/AAAA", max_chars=10, icon=":material/event:",
                     on_change=formato_fecha, args=(clave,))
    if not avisar(lugar, avisos, clave):
        lugar.caption(D.fecha_con_dia(calc[clave]) if calc[clave] else ayuda)


def monto(etiqueta: str, clave: str, lugar, avisos: dict, **extra) -> bool:
    lugar.text_input(etiqueta, key=clave, placeholder="0.00", icon=":material/attach_money:", on_change=dar_formato,
                     args=(clave,), **extra)
    return not avisar(lugar, avisos, clave)


def recibir_firma(archivo) -> None:
    ss = st.session_state
    huella = archivo.file_id if archivo is not None else None
    if huella == ss["_firma_huella"]:
        return
    ss["_firma_huella"], ss["_firma_error"] = huella, ""
    if archivo is None:
        ss["firma"] = ss["_firma_base"]
        return
    try:
        ss["firma"] = D.preparar_firma(archivo.getvalue())
    except ValueError as error:
        ss["firma"] = ss["_firma_base"]
        ss["_firma_error"] = str(error)


# ------------------------- Secciones de Captura --------------------------

def barra_hojas_trabajo(firmante: dict) -> None:
    ss = st.session_state
    hojas = ss["hojas"]
    hojas_listas = [
        (h["captura"], h.get("dias", 0))
        for h in hojas
        if any(not D.validar(h["captura"], firmante, doc) for doc in TITULOS)
    ]

    with st.container(key="sheets_bar"):
        ui.sheets_bar_header(
            total_hojas=len(hojas),
            listas_n=len(hojas_listas),
            indice_actual=indice_hoja_activa(),
        )

        # Pestañas de hojas en bloques de hasta 4 por renglón
        por_fila = 4
        for inicio in range(0, len(hojas), por_fila):
            tramo = hojas[inicio : inicio + por_fila]
            cols = st.columns(max(len(tramo), 2) if len(hojas) == 1 else len(tramo))
            for col, h in zip(cols, tramo):
                num = hojas.index(h) + 1
                cap = h["captura"]
                nom = D.may(cap.get("nombre")) or "Nueva baja"
                if len(nom) > 17:
                    nom = nom[:15].rstrip() + "…"
                tiene_err = bool(D.mal_escritos(cap))
                lista = any(not D.validar(cap, firmante, doc) for doc in TITULOS)
                if tiene_err:
                    ico = ":material/error:"
                elif lista:
                    ico = ":material/check_circle:"
                elif hay_captura(cap):
                    ico = ":material/edit_note:"
                else:
                    ico = ":material/draft:"
                es_activa = h["id"] == ss["hoja_activa_id"]
                col.button(
                    f"Hoja {num} · {nom}",
                    key=f"tab_hoja_{h['id']}",
                    type="primary" if es_activa else "secondary",
                    icon=ico,
                    on_click=seleccionar_hoja,
                    args=(h["id"],),
                    width="stretch",
                )

        acc1, acc2, acc3, acc4 = st.columns(4)
        acc1.button("Nueva hoja", key="btn_nueva_hoja", icon=":material/add:",
                    help="Abre una hoja vacía para otra baja", on_click=crear_hoja_vacia, width="stretch")
        acc2.button("Misma empresa", key="btn_misma_emp", icon=":material/domain_add:",
                    help="Abre una hoja nueva copiando empresa, representante, sucursal, estado y ciudad",
                    on_click=crear_hoja_misma_empresa, width="stretch")
        acc3.button("Duplicar hoja", key="btn_dup_hoja", icon=":material/content_copy:",
                    help="Clona todos los datos de la hoja actual", on_click=duplicar_hoja_actual, width="stretch")
        acc4.button("Cerrar hoja", key="btn_cerrar_hoja", icon=":material/close:",
                    on_click=cerrar_hoja_actual, disabled=len(hojas) <= 1, width="stretch")

        if len(hojas_listas) >= 2:
            b_col1, b_col2 = st.columns([2.2, 1.1], vertical_alignment="center")
            with b_col1:
                ui.batch_ready_banner(len(hojas_listas))
            with b_col2:
                b_col2.download_button(
                    f"Descargar lote ZIP ({len(hojas_listas)} hojas)",
                    lambda: D.paquete_lote_zip(hojas_listas, firmante),
                    f"Lote finiquitos {D.ahora_local():%Y-%m-%d}.zip",
                    "application/zip",
                    key="btn_zip_lote",
                    type="primary",
                    icon=":material/folder_zip:",
                    on_click="ignore",
                    width="stretch",
                )


def seccion_empresa(completo: bool) -> None:
    with st.container(key="sec_empresa"):
        ui.card_header(
            anchor="sec-empresa",
            num="01",
            titulo="Empresa y lugar",
            subtitulo="Razón social, representante legal y plaza donde se suscriben los documentos.",
            icono="building",
            color="blue",
            completo=completo,
        )
        a, b = st.columns([7, 5])
        a.selectbox("Empresa (razón social)", D.EMPRESAS, index=None, key="empresa", accept_new_options=True,
                    placeholder="Elige o escribe la razón social", on_change=completar_empresa)
        if st.session_state["empresa"] and st.session_state["empresa"] not in D.EMPRESAS:
            a.caption("No está en el catálogo: se imprime tal como la escribiste.")
        texto("Representante legal", "representante", b, placeholder="NOMBRE DEL APODERADO O REPRESENTANTE")
        a, b, c = st.columns([5, 4, 3])
        texto("Sucursal (nombre comercial)", "sucursal", a, placeholder="EJ. SUCURSAL CENTRO")
        b.selectbox("Estado", D.ESTADOS, index=None, key="estado", placeholder="Elige un estado")
        texto("Ciudad", "ciudad", c, placeholder="EJ. CULIACÁN")


def seccion_colaborador(calc: dict, avisos: dict, completo: bool, con_error: bool) -> None:
    ss = st.session_state
    with st.container(key="sec_colaborador"):
        ui.card_header(
            anchor="sec-colaborador",
            num="02",
            titulo="Colaborador",
            subtitulo="Identidad, vigencia laboral, sueldo mensual y campos configurables para documentos.",
            icono="user",
            color="indigo",
            completo=completo,
            con_error=con_error,
        )
        a, b = st.columns([8, 4])
        texto("Nombre completo", "nombre", a, placeholder="APELLIDOS Y NOMBRE(S) DEL COLABORADOR")
        texto("ID de nómina (opcional)", "idNomina", b, mayusculas=False, placeholder="Ej. 10482")

        a, b, c = st.columns(3)
        fecha("Fecha de ingreso", "ingreso", a, calc, avisos)
        fecha("Fecha de baja", "baja", b, calc, avisos, "Es la fecha que llevan los documentos.")
        if monto("Sueldo mensual", "sueldo", c, avisos):
            c.caption(md("Diario: " + D.pesos(calc["diario"])))

        activos_opt = sum(bool(ss.get(k, True)) for k in OPCIONALES)
        ui.optional_fields_bar(activos_opt)

        col_h, col_rfc, col_nss = st.columns(3)
        with col_h, st.container(key="opt_hora"):
            inc_hora = st.toggle("Incluir hora del convenio", key="incluirHora")
            texto("Hora del convenio", "hora", mayusculas=False, placeholder="09:30", max_chars=5,
                  icon=":material/schedule:", disabled=not inc_hora)
            if not inc_hora:
                st.caption("Omitida: el convenio no mostrará hora.")
            elif not avisar(st, avisos, "hora"):
                st.caption(f"Vacía, sale {D.hora_local(ss['estado'] or '')} hrs.")

        with col_rfc, st.container(key="opt_rfc"):
            inc_rfc = st.toggle("Incluir RFC", key="incluirRfc")
            texto("RFC (constancia)", "rfc", max_chars=13, placeholder="13 CARACTERES",
                  disabled=not inc_rfc)
            if not inc_rfc:
                st.caption("Omitido: no se pedirá ni imprimirá en la constancia.")
            else:
                avisar(st, avisos, "rfc")

        with col_nss, st.container(key="opt_nss"):
            inc_nss = st.toggle("Incluir Seguro Social (NSS)", key="incluirNss")
            texto("Número de seguro social", "nss", mayusculas=False, max_chars=11,
                  placeholder="11 dígitos sin guiones", disabled=not inc_nss)
            if not inc_nss:
                st.caption("Omitido: no se pedirá ni imprimirá en la constancia.")
            else:
                avisar(st, avisos, "nss")


def seccion_percepciones(calc: dict, avisos: dict, completo: bool, con_error: bool) -> None:
    anchos = [5.2, 1.8, 1.8, 3.2]
    anio = str(calc["baja"].year) if calc["baja"] else "Año"
    with st.container(key="sec_percepciones"):
        ui.card_header(
            anchor="sec-percepciones",
            num="03",
            titulo="Percepciones del finiquito",
            subtitulo="Copia los días y montos del cálculo de nómina. Deja vacío lo que no aplique.",
            icono="coins",
            color="teal",
            completo=completo,
            con_error=con_error,
        )
        cabecera = st.columns(anchos)
        for columna, (rotulo, der) in zip(
            cabecera,
            (("Concepto", False), ("Año", False), ("Días", False), ("Monto", True)),
        ):
            with columna:
                ui.perc_col_header(rotulo, derecha=der)

        for p in D.PERCEPCIONES:
            nombre, de, clave = p["nombre"], " de " + p["nombre"].lower(), f"perc.{p['k']}."
            fila = st.columns(anchos, vertical_alignment="center")
            with fila[0]:
                ui.perc_row_label(nombre, obligatorio=bool(p.get("obligatorio")), negrita=bool(p.get("negrita")))
            if p.get("periodo"):
                fila[1].text_input("Año" + de, key=clave + "periodo", label_visibility="collapsed", placeholder=anio,
                                   max_chars=4)
                avisar(fila[1], avisos, clave + "periodo")
            if p.get("dias"):
                fila[2].text_input("Días" + de, key=clave + "dias", label_visibility="collapsed", placeholder="Días")
                avisar(fila[2], avisos, clave + "dias")
            monto("Monto" + de, clave + "monto", fila[3], avisos, label_visibility="collapsed")

        ui.perc_total_bar(D.pesos(calc["percepciones"]))


def seccion_descuentos(calc: dict, avisos: dict) -> None:
    if any(clave.startswith("desc.") for clave in avisos):
        resumen = ":red[revisa los montos]"
    else:
        resumen = D.pesos(calc["deducciones"]) if calc["deducciones"] else "sin descuentos"
    with st.container(key="sec_descuentos"):
        with st.expander(md(f"**4. Descuentos y deducciones:** {resumen}"), key="exp_descuentos",
                         icon=":material/remove_circle_outline:"):
            columnas = st.columns(2)
            for i, d in enumerate(D.DESCUENTOS):
                monto(d["nombre"], f"desc.{d['k']}", columnas[i % 2], avisos)


def seccion_firmante() -> None:
    ss = st.session_state
    recibir_firma(ss.get("f_archivo"))
    if not D.limpio(ss["f_nombre"]):
        estado = "pendiente"
    else:
        estado = D.may(ss["f_nombre"]) + ("" if ss.get("firma") else ", sin firma")
    with st.container(key="sec_firmante"):
        with st.expander(f"**Quién firma la constancia laboral:** {estado}", key="exp_firmante",
                         icon=":material/draw:"):
            st.caption("Solo se usa para generar la constancia laboral membretada.")
            a, b = st.columns(2)
            texto("Nombre de quien firma", "f_nombre", a, placeholder="LIC. NOMBRE Y APELLIDOS")
            texto("Área", "f_area", b)
            a.text_input("Teléfono de contacto", key="f_telefono", icon=":material/call:")
            b.text_input("Correo electrónico", key="f_correo", icon=":material/mail:")
            st.text_area("Dirección de la oficina (sale arriba a la derecha, un renglón por línea)", key="f_direccion",
                         height=100, placeholder="Calle y número\nColonia y código postal\nCiudad, estado")
            if ss.get("firma"):
                a, b = st.columns([3, 2], vertical_alignment="center")
                a.file_uploader("Firma (foto o escaneo en fondo blanco)", type=["png", "jpg", "jpeg", "webp", "bmp"],
                                key="f_archivo", max_upload_size=5)
                if ss["_firma_error"]:
                    a.error(ss["_firma_error"])
                b.image(ss["firma"]["png"], caption="Vista previa de tinta en la constancia", width=240)
            else:
                st.file_uploader("Firma (foto o escaneo en fondo blanco)", type=["png", "jpg", "jpeg", "webp", "bmp"],
                                 key="f_archivo", max_upload_size=5)
                if ss["_firma_error"]:
                    st.error(ss["_firma_error"])


# -------------------- Vista 2: Importar y Editar PDF ---------------------

def vista_importar_pdf() -> None:
    with st.container(key="import_portal"):
        ui.pdf_import_header()
        archivo = st.file_uploader(
            "PDF a editar",
            type=["pdf"],
            key="pdf_para_editar",
            label_visibility="collapsed",
        )
        if archivo is None:
            st.caption("Sube cualquier PDF generado por esta herramienta (o convenios/constancias anteriores en PDF con texto) para prellenar una hoja automáticamente.")
            return

        try:
            extraido = D.leer_pdf_documento(archivo.getvalue())
        except ValueError as err:
            st.error(str(err))
            return

        cap = extraido["captura"]
        dias_det = extraido["dias"]
        calc_det = D.calcular(cap)
        doc_key = extraido["documento"]

        plaza_partes = [x for x in (cap.get("sucursal"), cap.get("ciudad"), cap.get("estado")) if D.limpio(x)]
        plaza_txt = " · ".join(plaza_partes) if plaza_partes else "—"
        vig_txt = (
            f"{cap['ingreso']} → {cap['baja']}" if cap.get("ingreso") and cap.get("baja")
            else (cap.get("baja") or cap.get("ingreso") or "—")
        )
        sueldo_txt = D.pesos(calc_det["sueldo"]) if calc_det["sueldo"] else "No incluido en este documento"
        n_op = 0 if dias_det == "todos" else int(dias_det)
        esc_txt = (
            f"{ROTULOS_DIAS.get(dias_det, 'Sin días')} · {D.pesos(calc_det['opciones'][n_op]['total'])}"
            if doc_key == "convenio" else TITULOS.get(doc_key, "Documento")
        )

        ui.pdf_import_preview(
            tipo_titulo=TITULOS.get(doc_key, "Documento").upper(),
            origen=extraido["origen"],
            colaborador=cap.get("nombre", ""),
            empresa=cap.get("empresa", ""),
            plaza=plaza_txt,
            vigencia=vig_txt,
            sueldo_txt=sueldo_txt,
            escenario_txt=esc_txt,
        )

        col_btn1, col_btn2 = st.columns([2, 1.2], vertical_alignment="center")
        col_btn1.button(
            "Abrir en una nueva hoja para editar datos",
            key="btn_cargar_pdf_hoja",
            type="primary",
            icon=":material/edit_document:",
            on_click=abrir_captura_en_nueva_hoja,
            args=(cap, dias_det, f"PDF cargado en hoja nueva: {cap.get('nombre') or 'Listo para editar'}"),
            width="stretch",
        )
        col_btn2.caption("Se abrirá en tu mesa de hojas con todos los campos cargados para que ajustes lo que necesites.")


# --------------------- Vista 3: Centro de Historial ----------------------

def vista_historial(firmante: dict) -> tuple[str, dict] | None:
    ss = st.session_state
    historial = ss["historial"]
    pedido_hist = None

    empresas_set = {D.may(r["captura"].get("empresa")) for r in historial if D.limpio(r["captura"].get("empresa"))}
    suma_base = sum(D.calcular(r["captura"])["opciones"][0]["total"] for r in historial)

    ui.history_kpis(
        total_registros=len(historial),
        empresas_unicas=len(empresas_set),
        suma_base_txt=D.pesos(suma_base),
        hojas_abiertas=len(ss["hojas"]),
    )

    with st.container(key="hist_portal"):
        ui.card_header(
            anchor="sec-historial",
            num="HISTORIAL",
            titulo="Bitácora de convenios y exportación a Excel",
            subtitulo="Consulta, filtra, reabre cualquier convenio en una hoja nueva para corregirlo o descarga el libro Histórico oficial.",
            icono="history",
            color="teal",
            completo=bool(historial),
        )

        b1, b2, b3 = st.columns([2.0, 1.45, 1.05], vertical_alignment="bottom")
        busqueda = b1.text_input(
            "Buscar en el historial",
            key="hist_busqueda",
            placeholder="Filtra por nombre del colaborador, empresa o sucursal…",
            icon=":material/search:",
        )
        b2.download_button(
            "Descargar Excel",
            lambda: D.libro_historial([r["captura"] for r in historial]),
            f"Historico convenios {D.ahora_local():%Y-%m-%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_descargar_excel_hist",
            type="primary",
            icon=":material/table_view:",
            disabled=not historial,
            on_click="ignore",
            width="stretch",
        )
        if b3.button("Vaciar", key="btn_vaciar_hist", icon=":material/delete_sweep:",
                     help="Borra los registros del historial actual",
                     disabled=not historial, width="stretch"):
            ss["historial"] = []
            sincronizar_historial_servidor(reemplazar_todo=True)
            st.rerun()

        if not historial:
            ui.history_empty_state()
        else:
            termino = D.sin_acentos(busqueda)
            filtrados = [
                r for r in historial
                if not termino or any(
                    termino in D.sin_acentos(r["captura"].get(campo, ""))
                    for campo in ("nombre", "empresa", "sucursal", "ciudad", "idNomina")
                )
            ]

            if not filtrados:
                st.info("Ningún registro coincide con tu búsqueda.")
            else:
                tabla = []
                for r in filtrados:
                    cap = r["captura"]
                    c = D.calcular(cap)
                    tabla.append({
                        "Colaborador": D.may(cap.get("nombre")) or "SIN NOMBRE",
                        "ID": D.limpio(cap.get("idNomina")) or "—",
                        "Empresa": D.may(cap.get("empresa")) or "—",
                        "Sucursal": D.may(cap.get("sucursal")) or "—",
                        "Baja": D.fecha_corta(c["baja"]) if c["baja"] else "—",
                        "0 días": D.pesos(c["opciones"][0]["total"]),
                        "15 días": D.pesos(c["opciones"][15]["total"]),
                        "30 días": D.pesos(c["opciones"][30]["total"]),
                        "45 días": D.pesos(c["opciones"][45]["total"]),
                    })
                st.dataframe(tabla, hide_index=True, width="stretch")

                nombres = {
                    r["id"]: f"{i}. {D.may(r['captura'].get('nombre')) or 'Sin nombre'} — {D.may(r['captura'].get('sucursal')) or D.may(r['captura'].get('empresa'))}"
                    for i, r in enumerate(filtrados, start=1)
                }
                ac1, ac2, ac3, ac4 = st.columns([2.8, 1.6, 1.5, 1.1], vertical_alignment="bottom")
                elegido_id = ac1.selectbox("Registro seleccionado", list(nombres), key="_elegido_hist", format_func=nombres.get)
                reg_sel = next((r for r in filtrados if r["id"] == elegido_id), None)

                if reg_sel:
                    ac2.button(
                        "Editar en Hoja",
                        key="btn_hist_a_hoja",
                        type="primary",
                        icon=":material/open_in_new:",
                        on_click=abrir_captura_en_nueva_hoja,
                        args=(copy.deepcopy(reg_sel["captura"]), 0, "Registro abierto en una nueva hoja."),
                        width="stretch",
                    )
                    puede_pdf = not D.validar(reg_sel["captura"], firmante, "convenio")
                    if ac3.button("Ver PDF", key="btn_hist_pdf", icon=":material/picture_as_pdf:",
                                  disabled=not puede_pdf, width="stretch"):
                        pedido_hist = ("convenio", reg_sel["captura"])
                    if ac4.button("Quitar", key="btn_hist_borrar_uno", icon=":material/delete:", width="stretch"):
                        ss["historial"] = [r for r in ss["historial"] if r["id"] != elegido_id]
                        sincronizar_historial_servidor(reemplazar_todo=True)
                        st.rerun()

        with st.container(key="excel_restore"):
            with st.expander("Importar / restaurar registros desde un archivo Excel Histórico (.xlsx)",
                             icon=":material/upload_file:"):
                arch_xl = st.file_uploader("Archivo Excel Histórico", type=["xlsx"], key="xl_restaurar",
                                           label_visibility="collapsed")
                if arch_xl is not None:
                    try:
                        recuperadas = D.leer_excel_historial(arch_xl.getvalue())
                    except ValueError as err:
                        st.error(str(err))
                    else:
                        st.caption(f"Se detectaron **{len(recuperadas)}** colaboradores en el archivo Excel.")
                        if recuperadas and st.button("Agregar estos registros al historial", key="btn_confirmar_xl",
                                                     type="primary", icon=":material/playlist_add:"):
                            existentes = {(D.may(r["captura"].get("nombre")), str(r["captura"].get("baja")))
                                          for r in ss["historial"]}
                            nuevos = 0
                            for cap_xl in recuperadas:
                                huella_xl = (D.may(cap_xl.get("nombre")), str(cap_xl.get("baja")))
                                if huella_xl not in existentes:
                                    ss["historial"].append({"id": uuid.uuid4().hex, "captura": cap_xl})
                                    existentes.add(huella_xl)
                                    nuevos += 1
                            guardar_estado_disco()
                            st.toast(f"Se integraron {nuevos} registros al historial.")
                            st.rerun()

    return pedido_hist


# ----------------------- Total y documentos (barra lateral) -----------------------

def barra_lateral(captura: dict, firmante: dict, calc: dict, docs_listos: int) -> str | None:
    ss = st.session_state
    pedido = None
    tiene_datos = hay_captura(captura)
    h_act = hoja_activa()

    with st.sidebar:
        ui.sidebar_brand(
            num_hoja=indice_hoja_activa(),
            total_hojas=len(ss["hojas"]),
            colaborador=D.may(captura.get("nombre")),
        )
        n = 0 if ss["dias"] == "todos" else ss["dias"]
        opcion = calc["opciones"][n]
        if ss["dias"] == "todos":
            rotulo = "Base finiquito (4 hojas)"
        elif n:
            rotulo = f"Total con {n} días"
        else:
            rotulo = "Total del finiquito"

        ui.sidebar_vault(
            rotulo=rotulo,
            total_txt=D.pesos(opcion["total"]),
            letras_txt=D.num_letras(opcion["total"]),
            percepciones_txt=D.pesos(opcion["percepciones"]),
            deducciones_txt=D.pesos(opcion["deducciones"]),
            dias=ss["dias"],
            es_negativo=opcion["total"] < 0,
            activo=tiene_datos and opcion["total"] > 0,
        )

        totales = [md(D.pesos(calc["opciones"][d]["total"])) for d in D.DIAS_CONVENIO] + ["Un PDF de 4 hojas"]
        st.radio("Días de convenio", OPCIONES_DIAS, key="dias", format_func=ROTULOS_DIAS.get, captions=totales)
        h_act["dias"] = ss["dias"]

        ui.sidebar_section_label("DOCUMENTOS EN PDF", f"{docs_listos}/3 LISTOS")
        if not tiene_datos:
            st.caption("Llena la hoja activa para habilitar las descargas.")
        for documento in TITULOS:
            if boton_documento(documento, captura, firmante, tiene_datos):
                pedido = documento

        s_col1, s_col2 = st.columns(2)
        s_col1.button("Guardar", key="btn_side_guardar_hist", on_click=guardar_hoja_actual_en_historial,
                      help="Guarda esta hoja en el historial", icon=":material/bookmark_add:",
                      disabled=not tiene_datos, width="stretch")
        s_col2.button("Limpiar", key="btn_side_limpiar", on_click=limpiar_hoja_actual,
                      help="Limpia los campos de la hoja actual", icon=":material/restart_alt:", width="stretch")
        ui.sidebar_privacy_note()
    return pedido


def boton_documento(documento: str, captura: dict, firmante: dict, tiene_datos: bool) -> bool:
    errores = D.validar(captura, firmante, documento)
    if not errores:
        return st.button(BOTONES[documento], key=f"abrir_{documento}", type="primary",
                         icon=":material/picture_as_pdf:", width="stretch")
    ui.sidebar_doc_pending(
        titulo=TITULOS[documento],
        errores=errores,
        mostrar_detalle=tiene_datos,
    )
    return False


# ------------------------- Ventana del documento -------------------------

def abrir_documento(documento: str, captura: dict, firmante: dict, id_registro: str | None = None, dias_override=None) -> None:
    dias = st.session_state["dias"] if dias_override is None else dias_override
    pdf = D.pdf_documento(documento, captura, firmante, dias)
    nombre = D.nombre_archivo(documento, captura, dias)
    titulo = TITULOS[documento]
    if documento == "convenio":
        titulo += ", " + DETALLE_DIAS[dias]
        if id_registro:
            anotar_en_historial(captura, id_registro)
    ventana = st.dialog(titulo, width="large", on_dismiss="rerun")(vista_previa)
    ventana(pdf, nombre)


def vista_previa(pdf: bytes, nombre: str) -> None:
    a, b = st.columns([2.2, 4.8], vertical_alignment="center")
    a.download_button("Descargar PDF", pdf, nombre, "application/pdf", key="descargar", type="primary",
                      icon=":material/download:", on_click="ignore", width="stretch")
    with b:
        ui.dialog_file_badge(nombre)
    try:
        st.pdf(pdf, height=700)
    except Exception:
        b64 = base64.b64encode(pdf).decode("ascii")
        ui.put(
            f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="680" '
            f'style="border:1px solid rgba(148,163,184,0.22);border-radius:14px;background:#0B1326;" '
            f'title="Vista previa PDF"></iframe>'
        )


# -------------------------------- Página ---------------------------------

pedir_clave()
iniciar_estado()

ss = st.session_state
if ss.get("_hidratar_pantalla") or "dias" not in ss:
    hidratar_widgets_desde_estado()
    ss["_hidratar_pantalla"] = False

if ss.get("_toast_pendiente"):
    st.toast(ss["_toast_pendiente"])
    ss["_toast_pendiente"] = ""

topbar_slot = st.empty()
ui.hero()

# Selector principal de modo de trabajo (3 módulos del estudio)
with st.container(key="mode_dock"):
    m1, m2, m3 = st.columns(3)
    vista = ss["vista_activa"]
    m1.button(
        f"Hojas de captura ({len(ss['hojas'])})",
        key="nav_modo_captura",
        type="primary" if vista == "captura" else "secondary",
        icon=":material/tab:",
        on_click=cambiar_vista,
        args=("captura",),
        width="stretch",
    )
    m2.button(
        "Subir y editar PDF",
        key="nav_modo_importar",
        type="primary" if vista == "importar" else "secondary",
        icon=":material/upload_file:",
        on_click=cambiar_vista,
        args=("importar",),
        width="stretch",
    )
    m3.button(
        f"Historial y Excel ({len(ss['historial'])})",
        key="nav_modo_historial",
        type="primary" if vista == "historial" else "secondary",
        icon=":material/history:",
        on_click=cambiar_vista,
        args=("historial",),
        width="stretch",
    )

pedido_extra = None

if vista == "captura":
    if "nombre" not in ss:
        hidratar_widgets_desde_estado()

    captura_prev = leer_captura_widgets()
    calc = D.calcular(captura_prev)
    avisos = D.mal_escritos(captura_prev)

    emp_ok = all(D.limpio(captura_prev.get(k)) for k in ("empresa", "representante", "sucursal", "estado", "ciudad"))
    col_err = any(k in avisos for k in ("ingreso", "baja", "sueldo", "hora", "rfc", "nss"))
    col_ok = (
        not col_err
        and all(D.limpio(captura_prev.get(k)) for k in ("nombre", "ingreso", "baja", "sueldo"))
        and bool(calc["ingreso"] and calc["baja"] and calc["sueldo"])
    )
    perc_err = any(k.startswith("perc.") for k in avisos)
    obligatorias_llenas = sum(
        1
        for p in D.PERCEPCIONES
        if p.get("obligatorio")
        and D.es_monto(captura_prev["perc"][p["k"]]["monto"])
        and D.a_centavos(captura_prev["perc"][p["k"]]["monto"]) is not None
    )
    perc_ok = not perc_err and obligatorias_llenas == sum(1 for p in D.PERCEPCIONES if p.get("obligatorio"))

    sheets_slot = st.empty()
    workflow_slot = st.empty()
    kpi_slot = st.empty()

    seccion_empresa(emp_ok)
    seccion_colaborador(calc, avisos, col_ok, col_err)
    seccion_percepciones(calc, avisos, perc_ok, perc_err)

    ui.section_divider(
        "sec-complementos",
        "Descuentos y firmante de constancia",
        "Deducciones aplicables al finiquito y datos de quien suscribe la constancia laboral.",
        "layers",
    )
    seccion_descuentos(calc, avisos)
    seccion_firmante()

    # Sincroniza lo último capturado en pantalla hacia la hoja activa y disco
    respaldar_widgets_en_hoja_activa()
    guardar_estado_disco()
    captura = hoja_activa()["captura"]
    firmante = leer_firmante()
    docs_listos = sum(1 for doc in TITULOS if not D.validar(captura, firmante, doc))

    with sheets_slot.container():
        barra_hojas_trabajo(firmante)

    dias_sel = ss.get("dias", 0)
    n_activo = 0 if dias_sel == "todos" else dias_sel
    opcion_activa = calc["opciones"][n_activo]
    if dias_sel == "todos":
        etiqueta_dias = "Paquete de 4 hojas (0, 15, 30 y 45d)"
    elif n_activo == 0:
        etiqueta_dias = "Sin días extra · Transferencia"
    else:
        etiqueta_dias = f"+{n_activo} días ({D.pesos(opcion_activa['extra'])}) · Efectivo"

    with workflow_slot.container():
        ui.workflow_rail(
            emp_ok=emp_ok,
            col_ok=col_ok,
            col_err=col_err,
            perc_ok=perc_ok,
            perc_err=perc_err,
            docs_listos=docs_listos,
        )
    with kpi_slot.container():
        ui.kpi_strip(
            percepciones_txt=D.pesos(calc["percepciones"]),
            obligatorias_llenas=obligatorias_llenas,
            deducciones_txt=D.pesos(calc["deducciones"]),
            deducciones_activas=sum(1 for d in calc["desc"] if d["monto"] > 0),
            diario_txt=D.pesos(calc["diario"]),
            antiguedad_txt=resumen_antiguedad(calc),
            total_activo_txt=D.pesos(opcion_activa["total"]),
            etiqueta_dias=etiqueta_dias,
            es_negativo=opcion_activa["total"] < 0,
        )

elif vista == "importar":
    ss["_hidratar_pantalla"] = True
    captura = hoja_activa()["captura"]
    calc = D.calcular(captura)
    avisos = D.mal_escritos(captura)
    firmante = leer_firmante()
    docs_listos = sum(1 for doc in TITULOS if not D.validar(captura, firmante, doc))
    vista_importar_pdf()

else:
    ss["_hidratar_pantalla"] = True
    captura = hoja_activa()["captura"]
    calc = D.calcular(captura)
    avisos = D.mal_escritos(captura)
    firmante = leer_firmante()
    docs_listos = sum(1 for doc in TITULOS if not D.validar(captura, firmante, doc))
    pedido_extra = vista_historial(firmante)

ui.footer()

with topbar_slot.container():
    ui.topbar(
        estado_actual=captura.get("estado") or "",
        hora_txt=D.hora_local(captura.get("estado") or ""),
        docs_listos=docs_listos,
        hay_datos=hay_captura(captura),
        num_avisos=len(avisos),
        historial_n=len(ss["historial"]),
    )

pedido_side = barra_lateral(captura, firmante, calc, docs_listos)
if pedido_side:
    abrir_documento(pedido_side, captura, firmante, id_registro=hoja_activa()["id"])
elif pedido_extra:
    doc_h, cap_h = pedido_extra
    abrir_documento(doc_h, cap_h, firmante, id_registro=None, dias_override=0)
