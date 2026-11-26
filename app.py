"""Convenios y finiquitos · app para Recursos Humanos.

Para probarla en tu equipo:   streamlit run app.py

Este archivo solo arma la pantalla. Los cálculos, los textos y los PDF están en documentos.py,
y el sistema visual Noir Liquid Aurora vive en ui.py y assets/styles.css.
"""
from __future__ import annotations

import base64
import hmac
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

TITULOS = {"convenio": "Convenio de finiquito", "renuncia": "Carta de renuncia", "constancia": "Constancia laboral"}
BOTONES = {"convenio": "Ver y descargar el convenio", "renuncia": "Ver y descargar la renuncia",
           "constancia": "Ver y descargar la constancia"}
OPCIONES_DIAS = [0, 15, 30, 45, "todos"]
# Los rótulos de las opciones deben ser fijos: Streamlit identifica la opción elegida por su texto.
ROTULOS_DIAS = {0: "Sin días", 15: "15 días", 30: "30 días", 45: "45 días", "todos": "Las cuatro opciones"}
DETALLE_DIAS = {0: "sin días de convenio", 15: "con 15 días", 30: "con 30 días", 45: "con 45 días",
                "todos": "las cuatro opciones"}


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


# ------------------------------- Estado ----------------------------------
# Cada campo de la pantalla guarda su valor en st.session_state con la misma clave que usa
# documentos.py para nombrarlo: "nombre", "sueldo", "perc.aguinaldo.monto", "desc.faltante"...

TEXTOS = ("representante", "sucursal", "ciudad", "nombre", "idNomina", "ingreso", "baja", "sueldo", "hora", "rfc", "nss")


def valores_iniciales() -> dict:
    """Clave de cada campo de la captura y su valor vacío."""
    v = {"empresa": None, "estado": None}
    v.update({k: "" for k in TEXTOS})
    for p in D.PERCEPCIONES:
        v.update({f"perc.{p['k']}.{campo}": "" for campo in ("periodo", "dias", "monto")})
    v.update({f"desc.{d['k']}": "" for d in D.DESCUENTOS})
    return v


def iniciar_estado() -> None:
    ss = st.session_state
    for k, v in valores_iniciales().items():
        ss.setdefault(k, v)
    if ss.get("dias") not in OPCIONES_DIAS:
        ss["dias"] = 0
    ss.setdefault("historial", [])          # convenios abiertos en esta sesión
    ss.setdefault("_id", uuid.uuid4().hex)  # identifica la captura actual dentro del historial
    if "f_nombre" not in ss:                # quién firma: valores por defecto tomados de los secretos
        f = dict(secreto("firmante", {}) or {})
        ss["f_nombre"] = D.may(f.get("nombre"))
        ss["f_area"] = D.may(f.get("area") or "RECURSOS HUMANOS")
        ss["f_telefono"] = str(f.get("telefono") or "")
        ss["f_correo"] = str(f.get("correo") or "")
        ss["f_direccion"] = str(f.get("direccion") or "")
        ss["firma"] = ss["_firma_base"] = firma_de_archivo(f.get("firma"))
        ss["_firma_huella"] = None
        ss["_firma_error"] = ""


def firma_de_archivo(ruta) -> dict | None:
    """Firma por defecto: ruta a una imagen indicada en los secretos (opcional)."""
    if not ruta:
        return None
    archivo = Path(__file__).resolve().parent / str(ruta)
    try:
        return D.preparar_firma(archivo.read_bytes())
    except (OSError, ValueError):
        return None


def leer_captura() -> dict:
    """Junta lo escrito en pantalla en el diccionario que entiende documentos.py."""
    ss = st.session_state
    s = D.captura_vacia()
    for k in ("empresa", "estado") + TEXTOS:
        s[k] = ss.get(k)
    for p in D.PERCEPCIONES:
        s["perc"][p["k"]] = {campo: ss.get(f"perc.{p['k']}.{campo}", "") for campo in ("periodo", "dias", "monto")}
    for d in D.DESCUENTOS:
        s["desc"][d["k"]] = ss.get(f"desc.{d['k']}", "")
    return s


def leer_firmante() -> dict:
    ss = st.session_state
    return {"nombre": ss["f_nombre"], "area": ss["f_area"], "telefono": ss["f_telefono"], "correo": ss["f_correo"],
            "direccion": ss["f_direccion"], "firma": ss.get("firma")}


def hay_captura(s: dict) -> bool:
    """True si ya se escribió algo en el formulario."""
    vacia = D.captura_vacia()
    sueltos = any(D.limpio(s[k]) for k in s if k not in ("perc", "desc"))
    return sueltos or s["perc"] != vacia["perc"] or s["desc"] != vacia["desc"]


def resumen_antiguedad(calc: dict) -> str:
    """Texto informativo de antigüedad o referencia del salario diario."""
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


# ----------------------------- Acciones ----------------------------------

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
        return  # mal escrito: se deja como está y se avisa junto al campo
    if c is not None:
        st.session_state[clave] = D.numero(c)


def formato_fecha(clave: str) -> None:
    """Al salir de una fecha la deja como 20/02/2026 (se puede escribir 20-2-26 o 20022026)."""
    f = D.a_fecha(st.session_state[clave])
    if f:
        st.session_state[clave] = D.fecha_corta(f)


def limpiar_formulario() -> None:
    for k, v in valores_iniciales().items():
        st.session_state[k] = v
    st.session_state["_id"] = uuid.uuid4().hex


def anotar_en_historial(captura: dict, identificador: str) -> None:
    """Guarda (o actualiza) la captura en el historial de la sesión cuando se abre su convenio."""
    historial = st.session_state["historial"]
    # Se cambia el contenido de la misma lista (no se crea otra): así "Descargar Excel" siempre lee la versión al día.
    historial[:] = [{"id": identificador, "captura": captura}] + [r for r in historial if r["id"] != identificador]


def abrir_registro() -> None:
    """Regresa al formulario el registro elegido en el historial."""
    ss = st.session_state
    registro = next((r for r in ss["historial"] if r["id"] == ss.get("_elegido")), None)
    if not registro:
        return
    s = registro["captura"]
    for k in TEXTOS:
        ss[k] = s.get(k) or ""
    ss["empresa"], ss["estado"] = s.get("empresa") or None, s.get("estado") or None
    for p in D.PERCEPCIONES:
        for campo in ("periodo", "dias", "monto"):
            ss[f"perc.{p['k']}.{campo}"] = s["perc"][p["k"]].get(campo, "")
    for d in D.DESCUENTOS:
        ss[f"desc.{d['k']}"] = s["desc"].get(d["k"], "")
    ss["_id"] = registro["id"]


# ------------------------------ Pantalla ---------------------------------
# Lo que se escribe llega siempre al servidor, esté bien o mal: si algo no se entiende se avisa en rojo
# junto al campo y el documento no se puede abrir hasta corregirlo (ver documentos.mal_escritos y validar).

def texto(etiqueta: str, clave: str, lugar=st, mayusculas: bool = True, **extra) -> None:
    """Campo de texto ligado al estado; por defecto convierte a mayúsculas al salir."""
    if mayusculas:
        extra.update(on_change=a_mayusculas, args=(clave,))
    lugar.text_input(etiqueta, key=clave, **extra)


def avisar(lugar, avisos: dict, clave: str) -> bool:
    """Si el campo quedó mal escrito lo dice en rojo debajo de él. Devuelve True cuando hubo aviso."""
    if clave not in avisos:
        return False
    lugar.caption(md(f":red[{avisos[clave]}]"))
    return True


def fecha(etiqueta: str, clave: str, lugar, calc: dict, avisos: dict, ayuda: str = "") -> None:
    """Campo de fecha escrita como día/mes/año. Debajo la repite con letra, o avisa si no se entiende."""
    lugar.text_input(etiqueta, key=clave, placeholder="DD/MM/AAAA", max_chars=10, icon=":material/event:",
                     on_change=formato_fecha, args=(clave,))
    if not avisar(lugar, avisos, clave):
        lugar.caption(D.fecha_con_dia(calc[clave]) if calc[clave] else ayuda)


def monto(etiqueta: str, clave: str, lugar, avisos: dict, **extra) -> bool:
    """Campo de dinero: al salir se le da formato. Devuelve False si quedó mal escrito."""
    lugar.text_input(etiqueta, key=clave, placeholder="0.00", icon=":material/attach_money:", on_change=dar_formato,
                     args=(clave,), **extra)
    return not avisar(lugar, avisos, clave)


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
    with st.container(key="sec_colaborador"):
        ui.card_header(
            anchor="sec-colaborador",
            num="02",
            titulo="Colaborador",
            subtitulo="Identidad, vigencia laboral, sueldo mensual y registros para la constancia.",
            icono="user",
            color="indigo",
            completo=completo,
            con_error=con_error,
        )
        a, b = st.columns([8, 4])
        texto("Nombre completo", "nombre", a, placeholder="APELLIDOS Y NOMBRE(S) DEL COLABORADOR")
        texto("ID de nómina (opcional)", "idNomina", b, mayusculas=False, placeholder="Ej. 10482")
        a, b, c, d = st.columns(4)
        fecha("Fecha de ingreso", "ingreso", a, calc, avisos)
        fecha("Fecha de baja", "baja", b, calc, avisos, "Es la fecha que llevan los documentos.")
        if monto("Sueldo mensual", "sueldo", c, avisos):
            c.caption(md("Diario: " + D.pesos(calc["diario"])))
        texto("Hora del convenio", "hora", d, mayusculas=False, placeholder="09:30", max_chars=5,
              icon=":material/schedule:")
        if not avisar(d, avisos, "hora"):
            d.caption(f"Opcional. Vacía, sale `{D.hora_local(st.session_state['estado'] or '')}`.")
        a, b = st.columns(2)
        texto("RFC (para la constancia)", "rfc", a, max_chars=13, placeholder="13 CARACTERES CON HOMOCLAVE")
        avisar(a, avisos, "rfc")
        texto("Número de seguro social (para la constancia)", "nss", b, mayusculas=False, max_chars=11,
              placeholder="11 dígitos sin guiones")
        avisar(b, avisos, "nss")


def seccion_percepciones(calc: dict, avisos: dict, completo: bool, con_error: bool) -> None:
    anchos = [5.2, 1.8, 1.8, 3.2]
    anio = str(calc["baja"].year) if calc["baja"] else "Año"  # si no se escribe, sale el año de la baja
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
    recibir_firma(ss.get("f_archivo"))  # antes de dibujar, para que el rótulo de la sección ya diga si hay firma
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


def recibir_firma(archivo) -> None:
    """Procesa la imagen subida una sola vez; si se quita, regresa la firma por defecto."""
    ss = st.session_state
    huella = archivo.file_id if archivo is not None else None  # cambia con cada archivo que se sube
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


def seccion_historial() -> None:
    historial = st.session_state["historial"]
    with st.container(key="sec_historial"):
        with st.expander(f"**Historial de esta sesión:** {len(historial)}", key="exp_historial",
                         icon=":material/history:"):
            st.caption("Cada convenio que abres para descargar queda anotado aquí mientras la página siga abierta; no se "
                       "guarda en el servidor. Descarga el Excel antes de cerrar: sale con el acomodo de la hoja "
                       "Historico, cuatro renglones por persona.")
            if not historial:
                return
            tabla = []
            for r in historial:
                c = D.calcular(r["captura"])
                tabla.append({"Colaborador": D.may(r["captura"]["nombre"]), "Empresa": D.may(r["captura"]["empresa"]),
                              "Sucursal": D.may(r["captura"]["sucursal"]), "Baja": c["baja"],
                              "Finiquito": c["opciones"][0]["total"] / 100})
            st.dataframe(tabla, hide_index=True, column_config={
                "Baja": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Finiquito": st.column_config.NumberColumn(format="dollar")})
            a, b, c = st.columns([3, 5, 3], vertical_alignment="bottom")
            # El libro se arma al pulsar el botón, con lo que haya en el historial en ese momento.
            a.download_button("Descargar Excel", lambda: D.libro_historial([r["captura"] for r in historial]),
                              f"Historico convenios {D.ahora_local():%Y-%m-%d}.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              icon=":material/table_view:", on_click="ignore", width="stretch")
            nombres = {r["id"]: f"{i}. {D.may(r['captura']['nombre']) or 'Sin nombre'}"
                       for i, r in enumerate(historial, start=1)}
            b.selectbox("Volver a abrir", list(nombres), key="_elegido", format_func=nombres.get)
            c.button("Abrir en el formulario", on_click=abrir_registro, icon=":material/open_in_new:", width="stretch")


# ----------------------- Total y documentos (barra lateral) -----------------------

def barra_lateral(captura: dict, firmante: dict, calc: dict, docs_listos: int) -> str | None:
    """Muestra el total, las opciones de días y los documentos. Devuelve el documento que se pidió abrir."""
    ss = st.session_state
    pedido = None
    tiene_datos = hay_captura(captura)
    with st.sidebar:
        ui.sidebar_brand()
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

        ui.sidebar_section_label("DOCUMENTOS EN PDF", f"{docs_listos}/3 LISTOS")
        if not tiene_datos:
            st.caption("Llena el formulario para habilitar las descargas.")
        for documento in TITULOS:
            if boton_documento(documento, captura, firmante, tiene_datos):
                pedido = documento

        st.button("Limpiar formulario", on_click=limpiar_formulario, icon=":material/restart_alt:", width="stretch")
        ui.sidebar_privacy_note()
    return pedido


def boton_documento(documento: str, captura: dict, firmante: dict, tiene_datos: bool) -> bool:
    """Botón que abre el documento o, si todavía no se puede armar, la tarjeta de lo que falta."""
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
# Los documentos no se descargan directo desde la barra lateral a propósito. El botón de la barra
# es un botón normal: al pulsarlo, Streamlit manda junto con el clic lo último que se escribió
# (aunque el cursor siga dentro de un campo) y el PDF se arma con eso. La descarga queda dentro de
# una ventana que tapa el formulario, así que lo que se ve es exactamente lo que se descarga.

def abrir_documento(documento: str, captura: dict, firmante: dict) -> None:
    """Arma el PDF una sola vez, con lo que hay en pantalla, y lo abre para revisarlo y descargarlo."""
    dias = st.session_state["dias"]
    pdf = D.pdf_documento(documento, captura, firmante, dias)
    nombre = D.nombre_archivo(documento, captura, dias)
    titulo = TITULOS[documento]
    if documento == "convenio":
        titulo += ", " + DETALLE_DIAS[dias]
        anotar_en_historial(captura, st.session_state["_id"])
    # Al cerrar la ventana la página se vuelve a dibujar (on_dismiss) para que el historial muestre lo anotado.
    ventana = st.dialog(titulo, width="large", on_dismiss="rerun")(vista_previa)
    ventana(pdf, nombre)


def vista_previa(pdf: bytes, nombre: str) -> None:
    """Contenido de la ventana: el botón de descarga y el PDF a la vista."""
    a, b = st.columns([2.2, 4.8], vertical_alignment="center")
    a.download_button("Descargar PDF", pdf, nombre, "application/pdf", key="descargar", type="primary",
                      icon=":material/download:", on_click="ignore", width="stretch")
    with b:
        ui.dialog_file_badge(nombre)
    try:
        st.pdf(pdf, height=700)
    except Exception:  # si streamlit[pdf] no está instalado en local, incrusta el visor nativo del navegador
        b64 = base64.b64encode(pdf).decode("ascii")
        ui.put(
            f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="680" '
            f'style="border:1px solid rgba(148,163,184,0.22);border-radius:14px;background:#0B1326;" '
            f'title="Vista previa PDF"></iframe>'
        )


# -------------------------------- Página ---------------------------------

pedir_clave()
iniciar_estado()

captura = leer_captura()
calc = D.calcular(captura)
avisos = D.mal_escritos(captura)  # campos mal escritos: se avisa junto a cada uno

emp_ok = all(D.limpio(captura.get(k)) for k in ("empresa", "representante", "sucursal", "estado", "ciudad"))
col_err = any(k in avisos for k in ("ingreso", "baja", "sueldo", "hora", "rfc", "nss"))
col_ok = (
    not col_err
    and all(D.limpio(captura.get(k)) for k in ("nombre", "ingreso", "baja", "sueldo"))
    and bool(calc["ingreso"] and calc["baja"] and calc["sueldo"])
)
perc_err = any(k.startswith("perc.") for k in avisos)
obligatorias_llenas = sum(
    1
    for p in D.PERCEPCIONES
    if p.get("obligatorio")
    and D.es_monto(captura["perc"][p["k"]]["monto"])
    and D.a_centavos(captura["perc"][p["k"]]["monto"]) is not None
)
perc_ok = not perc_err and obligatorias_llenas == sum(1 for p in D.PERCEPCIONES if p.get("obligatorio"))

topbar_slot = st.empty()
ui.hero()
workflow_slot = st.empty()
kpi_slot = st.empty()

seccion_empresa(emp_ok)
seccion_colaborador(calc, avisos, col_ok, col_err)
seccion_percepciones(calc, avisos, perc_ok, perc_err)

ui.section_divider(
    "sec-complementos",
    "Complementos, constancia e historial",
    "Deducciones aplicables, firmante de la constancia laboral y bitácora exportable a Excel.",
    "layers",
)
seccion_descuentos(calc, avisos)
seccion_firmante()
seccion_historial()
ui.footer()

# La imagen de la firma se procesa al dibujar su sección, por eso quién firma se lee hasta este punto.
firmante = leer_firmante()
docs_listos = sum(1 for doc in TITULOS if not D.validar(captura, firmante, doc))

dias_sel = st.session_state["dias"]
n_activo = 0 if dias_sel == "todos" else dias_sel
opcion_activa = calc["opciones"][n_activo]
if dias_sel == "todos":
    etiqueta_dias = "Paquete de 4 hojas (0, 15, 30 y 45d)"
elif n_activo == 0:
    etiqueta_dias = "Sin días extra · Transferencia"
else:
    etiqueta_dias = f"+{n_activo} días ({D.pesos(opcion_activa['extra'])}) · Efectivo"

deducciones_activas = sum(1 for d in calc["desc"] if d["monto"] > 0)

with topbar_slot.container():
    ui.topbar(
        estado_actual=st.session_state.get("estado") or "",
        hora_txt=D.hora_local(st.session_state.get("estado") or ""),
        docs_listos=docs_listos,
        hay_datos=hay_captura(captura),
        num_avisos=len(avisos),
        historial_n=len(st.session_state["historial"]),
    )
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
        deducciones_activas=deducciones_activas,
        diario_txt=D.pesos(calc["diario"]),
        antiguedad_txt=resumen_antiguedad(calc),
        total_activo_txt=D.pesos(opcion_activa["total"]),
        etiqueta_dias=etiqueta_dias,
        es_negativo=opcion_activa["total"] < 0,
    )

pedido = barra_lateral(captura, firmante, calc, docs_listos)
if pedido:
    abrir_documento(pedido, captura, firmante)
