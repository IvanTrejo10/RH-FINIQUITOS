# -*- coding: utf-8 -*-
"""Sistema visual Noir Liquid Aurora · Edición Ejecutiva RH & Legal (Caprepa)."""
from __future__ import annotations

import base64
from functools import lru_cache
from html import escape
from pathlib import Path

import streamlit as st

ASSETS = Path(__file__).resolve().parent / "assets"


def _txt(valor) -> str:
    """Escapa HTML y protege el signo de pesos para que Streamlit no active fórmulas LaTeX."""
    return escape("" if valor is None else str(valor)).replace("$", "&#36;")


@lru_cache(maxsize=1)
def _logo_b64() -> str:
    ruta = ASSETS / "logo.png"
    if ruta.exists():
        return base64.b64encode(ruta.read_bytes()).decode("ascii")
    return ""


def icon(name: str, size: int = 20) -> str:
    paths = {
        "caprepa": (
            '<circle cx="14" cy="5.2" r="2.5" fill="currentColor" stroke="none"/>'
            '<circle cx="9.2" cy="9.8" r="2.6" fill="currentColor" stroke="none"/>'
            '<circle cx="15" cy="12.2" r="3" fill="currentColor" stroke="none"/>'
            '<circle cx="9.6" cy="16.8" r="3.4" fill="currentColor" stroke="none"/>'
            '<ellipse cx="12" cy="20.6" rx="8.5" ry="1.7" fill="currentColor" opacity="0.38" stroke="none"/>'
        ),
        "building": '<path d="M4 21h16M6 21V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v16"/><path d="M9 7h2M13 7h2M9 11h2M13 11h2M9 15h2M13 15h2"/>',
        "user": '<circle cx="12" cy="8" r="4"/><path d="M4 20a8 8 0 0 1 16 0"/>',
        "coins": '<ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v6c0 1.66 3.13 3 7 3s7-1.34 7-3V6"/><path d="M5 12v6c0 1.66 3.13 3 7 3s7-1.34 7-3v-6"/>',
        "discount": '<path d="M19 5 5 19"/><circle cx="7.5" cy="7.5" r="2.5"/><circle cx="16.5" cy="16.5" r="2.5"/>',
        "pen": '<path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
        "history": '<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5M12 7v5l3 2"/>',
        "file_pdf": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6M8 13h8M8 17h5"/>',
        "check": '<path d="m5 12.5 4.2 4.2L19 7"/>',
        "check_double": '<path d="m2.5 12.5 4 4L15 8"/><path d="m9.5 15.5 2 2L21 8"/>',
        "shield": '<path d="M12 3 4 6.5V12c0 5 8 9 8 9s8-4 8-9V6.5Z"/><path d="m8.8 12.2 2.2 2.3 4.5-5"/>',
        "spark": '<path d="m12 2.8 2.2 6.2L20.5 11l-6.3 2.1L12 19.5l-2.2-6.4L3.5 11l6.3-2Z"/><path d="M19 3v3M20.5 4.5h-3"/>',
        "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.2 2"/>',
        "calendar": '<rect x="3" y="4.5" width="18" height="16.5" rx="2.5"/><path d="M8 2.5v4M16 2.5v4M3 9.5h18"/>',
        "layers": '<path d="m12 3 9 4.8-9 4.8-9-4.8Z"/><path d="m3 12.5 9 4.8 9-4.8"/><path d="m3 17.2 9 4.8 9-4.8"/>',
        "zap": '<path d="M13 2 4.5 13H12l-1 9 8.5-11H12Z"/>',
        "bank": '<path d="m3 9 9-6 9 6H3Z"/><path d="M5 9v8M9 9v8M15 9v8M19 9v8M3 20h18"/>',
        "cash": '<rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/><path d="M6 12h.01M18 12h.01"/>',
        "alert": '<circle cx="12" cy="12" r="9"/><path d="M12 8v4.5M12 16h.01"/>',
        "lock": '<rect x="4" y="11" width="16" height="10" rx="2.5"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
        "briefcase": '<rect x="3" y="7" width="18" height="13" rx="2.5"/><path d="M8 7V5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M3 13h18"/>',
    }
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
        f'stroke="currentColor" stroke-width="1.85" stroke-linecap="round" '
        f'stroke-linejoin="round" aria-hidden="true">{paths[name]}</svg>'
    )


def put(markup: str) -> None:
    st.markdown(markup, unsafe_allow_html=True)


def apply_styles() -> None:
    css = (ASSETS / "styles.css").read_text(encoding="utf-8")
    put(f"<style>{css}</style>")


def auth_header() -> None:
    logo = _logo_b64()
    logo_html = (
        f'<div style="margin-bottom:14px"><span class="hero-logo-pill">'
        f'<img src="data:image/png;base64,{logo}" alt="Caprepa"/></span></div>'
        if logo else ""
    )
    put(
        '<div class="auth-stage anim-reveal">'
        f'<div class="auth-orb">{icon("lock", 26)}</div>'
        f'{logo_html}'
        '<h1>Convenios y finiquitos</h1>'
        '<p>Ingresa la clave de Recursos Humanos para abrir el estudio documental.</p>'
        '</div>'
    )


def topbar(
    *,
    estado_actual: str,
    hora_txt: str,
    docs_listos: int,
    hay_datos: bool,
    num_avisos: int,
    historial_n: int,
) -> None:
    zona_lbl = _txt(estado_actual or "HORA CENTRO")
    if len(zona_lbl) > 16:
        zona_lbl = zona_lbl[:14].rstrip() + "…"

    if num_avisos:
        state_cls = "warn"
        state_txt = f"Revisa {num_avisos} {'dato' if num_avisos == 1 else 'datos'}"
    elif docs_listos == 3:
        state_cls = "live"
        state_txt = "3/3 PDFs listos"
    elif docs_listos > 0:
        state_cls = "live"
        state_txt = f"{docs_listos}/3 PDFs listos"
    elif hay_datos:
        state_cls = "capturing"
        state_txt = "Captura en curso"
    else:
        state_cls = "idle"
        state_txt = "Estudio RH activo"

    hist_pill = (
        f'<a href="#sec-complementos" target="_self" class="nav-hist" style="text-decoration:none">'
        f'{icon("history", 13)} {historial_n} en sesión</a>'
        if historial_n else ""
    )

    put(
        '<nav class="floating-nav anim-reveal" style="--delay:0ms" aria-label="Barra principal">'
        '<div class="nav-brand">'
        f'<span class="nav-jewel">{icon("caprepa", 19)}</span>'
        '<span class="nav-wordmark">CAPREPA <b>RH Studio</b></span>'
        '</div>'
        '<div class="nav-links">'
        f'<a href="#sec-empresa" target="_self" class="nav-pill active">{icon("building", 13)}<span>Empresa</span></a>'
        f'<a href="#sec-colaborador" target="_self" class="nav-pill">{icon("user", 13)}<span>Colaborador</span></a>'
        f'<a href="#sec-percepciones" target="_self" class="nav-pill">{icon("coins", 13)}<span>Percepciones</span></a>'
        f'<a href="#sec-complementos" target="_self" class="nav-pill">{icon("layers", 13)}<span>Extras</span></a>'
        '</div>'
        '<div class="nav-right">'
        f'{hist_pill}'
        f'<span class="nav-clock">{icon("clock", 13)} {zona_lbl} · {_txt(hora_txt)}</span>'
        f'<span class="nav-status {state_cls}"><i class="live-dot"></i>{state_txt}</span>'
        '</div>'
        '</nav>'
    )


def hero() -> None:
    logo = _logo_b64()
    logo_pill = (
        f'<span class="hero-logo-pill"><img src="data:image/png;base64,{logo}" alt="Caprepa Servicios Financieros"/></span>'
        if logo else ""
    )
    put(
        '<header class="hero-stage anim-reveal" style="--delay:55ms">'
        '<div class="hero-aura aura-emerald" aria-hidden="true"></div>'
        '<div class="hero-aura aura-sapphire" aria-hidden="true"></div>'
        '<div class="hero-aura aura-gold" aria-hidden="true"></div>'
        '<svg class="hero-arcs" viewBox="0 0 1160 280" fill="none" aria-hidden="true" preserveAspectRatio="xMidYMid slice">'
        '<defs>'
        '<linearGradient id="rhArc1" x1="0%" y1="50%" x2="100%" y2="50%">'
        '<stop offset="0%" stop-color="#14B8A6" stop-opacity="0"/>'
        '<stop offset="28%" stop-color="#2DD4BF" stop-opacity="0.52"/>'
        '<stop offset="52%" stop-color="#FBBF24" stop-opacity="0.48"/>'
        '<stop offset="76%" stop-color="#38BDF8" stop-opacity="0.52"/>'
        '<stop offset="100%" stop-color="#818CF8" stop-opacity="0"/>'
        '</linearGradient>'
        '<linearGradient id="rhArc2" x1="0%" y1="50%" x2="100%" y2="50%">'
        '<stop offset="0%" stop-color="#38BDF8" stop-opacity="0"/>'
        '<stop offset="50%" stop-color="#FDE68A" stop-opacity="0.34"/>'
        '<stop offset="100%" stop-color="#38BDF8" stop-opacity="0"/>'
        '</linearGradient>'
        '</defs>'
        '<path class="arc-path arc-1" d="M 40 250 Q 580 -35 1120 250" stroke="url(#rhArc1)" stroke-width="1.4"/>'
        '<path class="arc-path arc-2" d="M 120 268 Q 580 28 1040 268" stroke="url(#rhArc2)" stroke-width="1" stroke-dasharray="5 9"/>'
        '<circle class="orbit-spark s1" cx="290" cy="146" r="3" fill="#5EEAD4"/>'
        '<circle class="orbit-spark s2" cx="870" cy="146" r="2.8" fill="#FBBF24"/>'
        '<circle class="orbit-spark s3" cx="580" cy="108" r="2.4" fill="#38BDF8"/>'
        '</svg>'
        '<div class="hero-inner">'
        '<div class="hero-top-badges">'
        f'{logo_pill}'
        '<span class="hero-kicker"><i class="kicker-dot"></i><span>RECURSOS HUMANOS · ESTUDIO DOCUMENTAL</span></span>'
        '</div>'
        '<h1 class="hero-title">Convenios y finiquitos<br><span class="hero-gradient"></span></h1>'
        '<p class="hero-sub">Captura los datos del colaborador en un solo flujo y descarga el convenio de finiquito, la carta de renuncia y la constancia laboral membretada en PDF.</p>'
        '<div class="hero-chips">'
        f'<span class="hero-chip gold">{icon("zap", 14)} Escenarios 0, 15, 30 y 45 días</span>'
        f'<span class="hero-chip">{icon("file_pdf", 14)} Convenio, Renuncia y Constancia</span>'
        f'<span class="hero-chip blue">{icon("history", 14)} Bitácora exportable a Excel</span>'
        f'<span class="hero-chip">{icon("shield", 14)} Memoria efímera segura</span>'
        '</div>'
        '</div>'
        '</header>'
    )


def workflow_rail(
    *,
    emp_ok: bool,
    col_ok: bool,
    col_err: bool,
    perc_ok: bool,
    perc_err: bool,
    docs_listos: int,
) -> None:
    s1 = "done" if emp_ok else "active"
    s2 = "warn" if col_err else ("done" if col_ok else ("active" if emp_ok else "pending"))
    s3 = "warn" if perc_err else ("done" if perc_ok else ("active" if (emp_ok and col_ok) else "pending"))
    s4 = "done" if docs_listos == 3 else ("active" if docs_listos > 0 else "pending")

    c1 = "lit" if emp_ok else "flow"
    c2 = "lit" if (emp_ok and col_ok and not col_err) else ("flow" if emp_ok else "dim")
    c3 = "lit" if docs_listos > 0 else ("flow" if perc_ok else "dim")

    def _node(num: str, state: str, ico: str, title: str, sub: str) -> str:
        glyph = icon("check", 16) if state == "done" else (icon("alert", 16) if state == "warn" else icon(ico, 16))
        return (
            f'<div class="flow-node {state}">'
            f'<span class="flow-ico">{glyph}</span>'
            '<div class="flow-txt">'
            f'<span class="flow-num">ETAPA {num}</span>'
            f'<strong>{_txt(title)}</strong>'
            f'<small>{_txt(sub)}</small>'
            '</div>'
            '</div>'
        )

    sub_docs = (
        "Los 3 PDFs habilitados" if docs_listos == 3
        else (f"{docs_listos} de 3 PDFs listos" if docs_listos else "Revisa la barra lateral")
    )

    put(
        '<div class="flow-rail anim-reveal" style="--delay:110ms">'
        + _node("01", s1, "building", "Empresa y sede", "Lista" if emp_ok else "Razón social y ciudad")
        + f'<div class="flow-conduit {c1}" aria-hidden="true"><span class="conduit-pulse"></span></div>'
        + _node("02", s2, "user", "Colaborador", "Revisa formato" if col_err else ("Datos completos" if col_ok else "Ingreso, baja y sueldo"))
        + f'<div class="flow-conduit {c2}" aria-hidden="true"><span class="conduit-pulse"></span></div>'
        + _node("03", s3, "coins", "Liquidación", "Revisa montos" if perc_err else ("Montos listos" if perc_ok else "Aguinaldo, vac. y prima"))
        + f'<div class="flow-conduit {c3}" aria-hidden="true"><span class="conduit-pulse"></span></div>'
        + _node("04", s4, "file_pdf", "Documentos PDF", sub_docs)
        + '</div>'
    )


def kpi_strip(
    *,
    percepciones_txt: str,
    obligatorias_llenas: int,
    deducciones_txt: str,
    deducciones_activas: int,
    diario_txt: str,
    antiguedad_txt: str,
    total_activo_txt: str,
    etiqueta_dias: str,
    es_negativo: bool,
) -> None:
    sub_desc = (
        f"{deducciones_activas} {'concepto aplicado' if deducciones_activas == 1 else 'conceptos aplicados'}"
        if deducciones_activas else "Sin descuentos capturados"
    )
    cls_total = "kpi-neg" if es_negativo else "kpi-gold"
    val_total_cls = "rose" if es_negativo else "gold"

    put(
        '<div class="kpi-grid anim-reveal" style="--delay:150ms">'
        '<div class="kpi-card">'
        f'<div class="kpi-top"><span>PERCEPCIONES BASE</span><i class="kpi-ico teal">{icon("coins", 15)}</i></div>'
        f'<strong class="kpi-val emerald">{_txt(percepciones_txt)}</strong>'
        f'<small>{obligatorias_llenas}/3 obligatorias capturadas</small>'
        '</div>'
        '<div class="kpi-card">'
        f'<div class="kpi-top"><span>DEDUCCIONES</span><i class="kpi-ico rose">{icon("discount", 15)}</i></div>'
        f'<strong class="kpi-val">{_txt(deducciones_txt)}</strong>'
        f'<small>{_txt(sub_desc)}</small>'
        '</div>'
        '<div class="kpi-card">'
        f'<div class="kpi-top"><span>SALARIO DIARIO</span><i class="kpi-ico blue">{icon("briefcase", 15)}</i></div>'
        f'<strong class="kpi-val">{_txt(diario_txt)}</strong>'
        f'<small>{_txt(antiguedad_txt)}</small>'
        '</div>'
        f'<div class="kpi-card {cls_total}">'
        f'<div class="kpi-top"><span>NETO ACTIVO</span><i class="kpi-ico gold">{icon("spark", 15)}</i></div>'
        f'<strong class="kpi-val {val_total_cls}">{_txt(total_activo_txt)}</strong>'
        f'<small>{_txt(etiqueta_dias)}</small>'
        '</div>'
        '</div>'
    )


def card_header(
    *,
    anchor: str,
    num: str,
    titulo: str,
    subtitulo: str,
    icono: str,
    color: str,
    completo: bool,
    con_error: bool = False,
) -> None:
    if con_error:
        pill_cls = "warn"
        pill_html = f'{icon("alert", 13)}<span>Revisa formato</span>'
    elif completo:
        pill_cls = "done"
        pill_html = f'{icon("check", 13)}<span>Completo</span>'
    else:
        pill_cls = "pending"
        pill_html = '<span>En captura</span>'

    put(
        f'<div class="card-head" id="{anchor}">'
        '<div class="card-head-left">'
        f'<div class="card-orb {color}">{icon(icono, 20)}</div>'
        '<div class="card-titles">'
        f'<span class="card-step-eyebrow">SECCIÓN {num}</span>'
        f'<h3>{_txt(titulo)}</h3>'
        f'<p>{_txt(subtitulo)}</p>'
        '</div>'
        '</div>'
        f'<span class="card-pill {pill_cls}">{pill_html}</span>'
        '</div>'
    )


def perc_col_header(texto: str, derecha: bool = False) -> None:
    cls = "perc-col-head right" if derecha else "perc-col-head"
    put(f'<div class="{cls}">{_txt(texto)}</div>')


def perc_row_label(nombre: str, obligatorio: bool = False, negrita: bool = False) -> None:
    cls = "perc-row-title bold" if negrita else "perc-row-title"
    extra = ""
    if obligatorio:
        extra = '<span class="req-badge">OBLIGATORIO *</span>'
    elif negrita:
        extra = '<span class="exenta-badge">+ DÍAS CONVENIO</span>'
    put(f'<div class="{cls}"><span>{_txt(nombre)}</span>{extra}</div>')


def perc_total_bar(total_pesos: str) -> None:
    put(
        '<div class="perc-total-bar">'
        '<div class="pt-left">'
        f'<span class="pt-ico">{icon("coins", 17)}</span>'
        '<div>'
        '<span class="pt-label">Total de percepciones base</span>'
        '<span class="pt-sub">Los rubros marcados con * son obligatorios para el convenio (escribe 0 si no aplica).</span>'
        '</div>'
        '</div>'
        f'<span class="pt-amount">{_txt(total_pesos)}</span>'
        '</div>'
    )


def section_divider(anchor: str, titulo: str, subtitulo: str, icono: str) -> None:
    put(
        f'<div class="section-header anim-reveal" id="{anchor}" style="--delay:190ms">'
        f'<span class="section-badge-ico">{icon(icono, 18)}</span>'
        '<div class="section-meta">'
        f'<h2>{_txt(titulo)}</h2>'
        f'<p>{_txt(subtitulo)}</p>'
        '</div>'
        '<span class="section-rule" aria-hidden="true"></span>'
        '</div>'
    )


def sidebar_brand() -> None:
    put(
        '<div class="side-brand">'
        f'<span class="side-brand-mark">{icon("caprepa", 21)}</span>'
        '<div>'
        '<div class="side-brand-title">RH<span> Finiquitos</span></div>'
        '<div class="side-brand-sub"><i class="live-dot"></i>BÓVEDA &amp; PDF STUDIO</div>'
        '</div>'
        '</div>'
    )


def sidebar_vault(
    *,
    rotulo: str,
    total_txt: str,
    letras_txt: str,
    percepciones_txt: str,
    deducciones_txt: str,
    dias,
    es_negativo: bool,
    activo: bool,
) -> None:
    if dias == "todos":
        modo_cls = "bank"
        modo_lbl = f'{icon("layers", 11)}<span>4 HOJAS</span>'
    elif dias == 0:
        modo_cls = "bank"
        modo_lbl = f'{icon("bank", 11)}<span>TRANSFERENCIA</span>'
    else:
        modo_cls = ""
        modo_lbl = f'{icon("cash", 11)}<span>EFECTIVO · {dias}D</span>'

    card_cls = "negative" if es_negativo else ("active" if activo else "")
    texto_inferior = "Los descuentos superan a las percepciones." if es_negativo else letras_txt

    put(
        f'<div class="vault-card {card_cls}">'
        '<div class="vault-head">'
        f'<span class="vault-eyebrow">{_txt(rotulo.upper())}</span>'
        f'<span class="vault-pay-badge {modo_cls}">{modo_lbl}</span>'
        '</div>'
        f'<div class="vault-amount">{_txt(total_txt)}</div>'
        f'<div class="vault-words">{_txt(texto_inferior)}</div>'
        '<div class="vault-splits">'
        f'<div class="vault-split pos"><span>+ Percepciones</span><b>{_txt(percepciones_txt)}</b></div>'
        f'<div class="vault-split neg"><span>− Deducciones</span><b>{_txt(deducciones_txt)}</b></div>'
        '</div>'
        '</div>'
    )


def sidebar_section_label(texto: str, derecha: str = "") -> None:
    der_html = f"<b>{_txt(derecha)}</b>" if derecha else ""
    put(f'<div class="side-section-label"><span>{_txt(texto)}</span>{der_html}</div>')


def sidebar_doc_pending(
    *,
    titulo: str,
    errores: list[tuple[str, str]],
    mostrar_detalle: bool,
) -> None:
    faltan = len(errores)
    badge_cls = "active" if mostrar_detalle else ""
    badge_txt = f"Faltan {faltan}" if mostrar_detalle else "En espera"

    lista_html = ""
    if mostrar_detalle and errores:
        visibles = [f"<li>{_txt(txt_err)}</li>" for _, txt_err in errores[:4]]
        if faltan > 4:
            visibles.append(f"<li>y {faltan - 4} datos más…</li>")
        lista_html = f'<ul class="doc-missing-list">{"".join(visibles)}</ul>'

    put(
        '<div class="doc-pending-card">'
        '<div class="doc-pending-top">'
        f'<span class="doc-pending-title"><span class="doc-pending-ico">{icon("file_pdf", 14)}</span>{_txt(titulo)}</span>'
        f'<span class="doc-pending-badge {badge_cls}">{_txt(badge_txt)}</span>'
        '</div>'
        f'{lista_html}'
        '</div>'
    )


def sidebar_privacy_note() -> None:
    put(
        '<div class="side-note">'
        f'<div class="side-note-head"><span class="shield-badge">{icon("shield", 15)}</span><h4>Privacidad en memoria</h4></div>'
        '<p>Lo que captures no se guarda en disco ni base de datos: arma los PDF al vuelo y se borra al cerrar la pestaña.</p>'
        '</div>'
    )


def dialog_file_badge(nombre_archivo: str) -> None:
    put(
        '<div class="pdf-meta-bar">'
        f'{icon("file_pdf", 17)}'
        f'<code title="{_txt(nombre_archivo)}">{_txt(nombre_archivo)}</code>'
        '</div>'
    )


def footer() -> None:
    put(
        '<footer class="studio-footer">'
        '<div class="footer-left"><strong>CAPREPA <span>·</span> Recursos Humanos</strong><span class="f-sep">·</span><span>Convenios, Renuncias y Constancias Laborales</span></div>'
        '<div>Cálculo exacto en centavos · Generación instantánea en PDF y Excel</div>'
        '</footer>'
    )
