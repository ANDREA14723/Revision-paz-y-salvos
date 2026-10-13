"""Revisión de paz y salvos (FO-AC-101) y cartas de continuidad (FO-AC-103) contra el FO-AC-83.

Ejecutar:  streamlit run app.py
"""
import html
import io

import pandas as pd

import streamlit as st

import validador as v

st.set_page_config(page_title="Revisión de paz y salvos", page_icon="🧾", layout="wide")
st.title("🧾 Revisión de paz y salvos y cartas de continuidad")
st.caption("FO-AC-101 / FO-AC-103 (PDF) contra el reporte de personal FO-AC-83 (Excel)")

# ------------------------------------------------------------------ barra lateral
with st.sidebar:
    st.header("Reglas")
    st.markdown("**Asunto esperado del FO-AC-101** (según tipo de vinculación y clase de riesgo del FO-AC-83)")
    asuntos = {}
    for k, etq in (("LABORAL", "LABORAL"), ("OPS_I_III", "OPS, riesgo I, II y III"), ("OPS_IV_V", "OPS, riesgo IV y V")):
        asuntos[k] = st.text_area(etq, v.CONFIG_DEFECTO["asuntos"][k], height=90)
    st.divider()
    una_pagina = st.checkbox("Cada página del PDF es un documento", value=True)
    solo_retiro = st.checkbox("Cruzar solo con filas RETIRO / INGRESO-RETIRO del FO-AC-83", value=True)
    usar_ocr = st.checkbox("Usar OCR en páginas escaneadas", value=True)
    tesseract = st.text_input("Ruta de tesseract.exe (solo Windows, si no está en el PATH)", "")

config = {
    "asuntos": asuntos,
    "solo_retiro": solo_retiro,
    "una_pagina_por_documento": una_pagina,
    "tesseract_cmd": tesseract.strip(),
}

# ------------------------------------------------------------------ carga de archivos
c1, c2 = st.columns(2)
pdfs = c1.file_uploader("PDF(s) con paz y salvos / cartas de continuidad", type="pdf",
                        accept_multiple_files=True)
xlsx = c2.file_uploader("FO-AC-83 reporte de personal (opcional, habilita el cruce)", type=["xlsx", "xlsm"])

if st.button("Revisar", type="primary", disabled=not pdfs):
    try:
        fo83 = v.cargar_fo83(io.BytesIO(xlsx.getvalue())) if xlsx else None
    except Exception as e:  # noqa: BLE001
        st.error(f"No pude leer el FO-AC-83: {e}")
        st.stop()
    archivos = [(f.name, f.getvalue()) for f in pdfs]
    with st.spinner("Leyendo PDFs (las páginas escaneadas tardan unos segundos cada una)…"):
        res = v.revisar(archivos, fo83, config, ocr=usar_ocr)
    st.session_state["res"] = res
    st.session_state["fo83"] = fo83
    st.session_state["pdf_bytes"] = dict(archivos)

res = st.session_state.get("res")
if not res:
    st.info("Sube el PDF (y el FO-AC-83 si quieres el cruce) y pulsa **Revisar**.")
    st.stop()

docs, faltantes, fo83 = res["docs"], res["faltantes"], st.session_state.get("fo83")
otros = res.get("otros", [])
for aviso in res.get("avisos", []):
    st.warning(aviso)

TIPOS = {  # clave: (emoji, etiqueta, color)
    "nombre": ("👤", "Nombre", "#B5524A"), "cedula": ("🪪", "Cédula", "#B5524A"),
    "cargo": ("💼", "Cargo", "#C98A1B"), "inicio": ("📅", "Fecha de ingreso", "#C98A1B"),
    "fin": ("📅", "Fecha de retiro", "#C98A1B"), "empresa": ("🏢", "Empresa", "#6B5BB5"),
    "asunto": ("📝", "Asunto / vinculación / riesgo", "#2F6FB5"), "seccion": ("🚫", "No está en RETIRO", "#B5524A"),
    "falta_excel": ("➖", "Falta en Excel", "#6B7280"), "falta_pdf": ("⚠️", "Falta en PDF", "#C98A1B"),
    "otras": ("🔵", "Otras observaciones", "#2F6FB5"),
}
st.markdown("""<style>
.card{border:1px solid #e5e7eb;border-radius:10px;padding:12px 14px;margin:10px 0;background:#fff}
.card-h{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;margin-bottom:8px}
.nm{font-weight:700}.cc{color:#6b7280;font-size:.9rem}
.badge{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.75rem;font-weight:600;margin:2px}
table.cmp{width:100%;border-collapse:collapse;font-size:.9rem}
table.cmp th{text-align:left;background:#f3f4f6;padding:6px}table.cmp td{padding:6px;border-top:1px solid #eee}
tr.bad td{background:#fdecea}tr.warn td{background:#fff6e0}
.msg{padding:8px;border-radius:8px;background:#fff6e0}.obs{font-size:.85rem;margin-top:6px}
.kpis{display:flex;gap:10px;flex-wrap:wrap}.kpi{flex:1;min-width:130px;border:1px solid #e5e7eb;border-radius:10px;padding:10px;text-align:center}
.kpi .n{font-size:1.6rem;font-weight:700}.kpi.bad .n{color:#B5524A}.kpi.ok .n{color:#2E7D4F}.kpi.warn .n{color:#C98A1B}
.banner{padding:12px;border-radius:10px;font-weight:700;margin-bottom:10px}.banner.ok{background:#e6f4ea}.banner.bad{background:#fdecea}
</style>""", unsafe_allow_html=True)

esc = lambda v: html.escape(str(v)) if str(v or "").strip() else "—"


def badges(tipos):
    return "".join(f'<span class="badge" style="background:{TIPOS[t][2]}1F;color:{TIPOS[t][2]};'
                   f'border:1px solid {TIPOS[t][2]}55">{TIPOS[t][0]} {TIPOS[t][1].upper()}'
                   f'{"" if t.startswith("falta") or t in ("seccion", "otras") else " NO COINCIDE"}</span>' for t in tipos)


def card(d):
    tipos = v.categorias_doc(d)
    cab = (f'<div class="card-h"><div><span class="nm">👤 {esc(d.nombre_firma or d.nombre)}</span><br>'
           f'<span class="cc">C.C. {esc(d.cedula_raw)} · pág. {esc(d.pagina)} · {esc(d.formato)} · lectura {esc(d.fuente)}'
           + (f' · emparejado con la fila {d.excel["fila_excel"]} del FO-AC-83'
              + (f' (la persona tiene {d.candidatas} filas)' if d.candidatas > 1 else '') if d.excel else '')
           + '</span></div>'
           f'<div>{badges(tipos)}</div></div>')
    filas = v.filas_comparacion(d)
    if filas:
        trs = ""
        for f in filas:
            clase = {"ERROR": "bad", "ALERTA": "warn"}.get(f["nivel"], "")
            res_ = {"ERROR": "✖ NO COINCIDE", "ALERTA": "⚠ REVISAR", "OK": "✔ Coincide"}[f["nivel"]]
            nota = f'<br><small>{esc(f["nota"])}</small>' if f["nota"] else ""
            trs += (f'<tr class="{clase}"><td>{esc(f["dato"])}</td><td>{esc(f["excel"])}</td>'
                    f'<td>{esc(f["pdf"])}</td><td>{res_}{nota}</td></tr>')
        cuerpo = ('<table class="cmp"><tr><th>Dato</th><th>📊 FO-AC-83 (Excel)</th><th>📄 Certificado (PDF)</th>'
                  f'<th>Resultado</th></tr>{trs}</table>')
    else:
        msgs = "<br>".join(esc(h[2]) for h in d.hallazgos if v.categoria(h[1]) in ("falta_excel", "seccion"))
        cuerpo = f'<div class="msg">⚠️ <b>No se pudo cruzar con el FO-AC-83.</b><br>{msgs}</div>'
    otras = "".join(f'<div class="obs">{"❌" if h[0] == v.ERROR else "⚠️"} {esc(h[2])}</div>' for h in v.otras_observaciones(d))
    return f'<div class="card">{cab}{cuerpo}{otras}</div>'


def card_falta_pdf(x):
    return (f'<div class="card"><div class="card-h"><div><span class="nm">👤 {esc(x["nombre"])}</span><br>'
            f'<span class="cc">C.C. {esc(x["cedula"])}</span></div><div>{badges(["falta_pdf"])}</div></div>'
            f'<div class="msg">⚠️ <b>Está en {esc(x["seccion"])} del FO-AC-83 pero NO hay paz y salvo ni carta de '
            f'continuidad para este cargo/periodo en los PDF.</b>'
            + ('<br><i>La persona sí entregó otro certificado, pero ninguno corresponde a esta fila.</i>' if x.get("entrego_otro") else '')
            + f'<br>Cargo: {esc(x["cargo"])} · Vinculación: {esc(x["tipo"])} · Riesgo: {esc(x["clase"])} · '
            f'Ingreso: {esc(v.fmt(x["ingreso"]))} · Retiro: {esc(v.fmt(x["retiro"]))} · Empresa: {esc(x["empresa"])}</div></div>')


# ------------------------------------------------------------------ resumen
con_problemas = [d for d in docs if d.estado != v.OK]
n_ok = len(docs) - len(con_problemas)
n_falta_xl = sum("falta_excel" in v.categorias_doc(d) for d in docs)
n_retiros = len([x for x in fo83["filas"] if "RETIRO" in x["seccion"]]) if fo83 else 0
cubiertos = max(n_retiros - len(faltantes), 0)

if not con_problemas and not faltantes:
    st.markdown('<div class="banner ok">✅ SIN OBSERVACIONES — todo coincide con el FO-AC-83</div>', unsafe_allow_html=True)
else:
    st.markdown(f'<div class="banner bad">🚨 {len(con_problemas)} certificado(s) con observaciones'
                f'{f" y {len(faltantes)} retiro(s) sin soporte" if fo83 else ""}</div>', unsafe_allow_html=True)

kpis = [("📄 Certificados en PDF", len(docs), ""), ("✅ Sin observaciones", n_ok, "ok"),
        ("✖ Con observaciones", len(con_problemas), "bad")]
if fo83:
    kpis += [("📊 Retiros en FO-AC-83", n_retiros, ""), ("⚠️ Retiros sin soporte", len(faltantes), "warn"),
             ("➖ Falta en Excel", n_falta_xl, "")]
st.markdown('<div class="kpis">' + "".join(f'<div class="kpi {c}"><div class="n">{n}</div><div class="l">{l}</div></div>'
                                           for l, n, c in kpis) + "</div>", unsafe_allow_html=True)
if fo83:
    p = fo83["periodo"]
    st.caption(f"FO-AC-83: contrato {fo83['contrato']} · {fo83['contratista']} · periodo {v.fmt(p[0])} a {v.fmt(p[1])}. "
               f"Solo se cruzan las filas RETIRO / INGRESO-RETIRO. Cobertura de retiros: {cubiertos} de {n_retiros}.")
else:
    st.warning("No cargaste el FO-AC-83: solo se muestran las revisiones internas de cada documento, sin comparación.")

if otros:
    with st.expander(f"📎 {len(otros)} página(s) de otros soportes ignoradas (transferencias, comprobantes, etc.)"):
        st.caption("No son FO-AC-101 ni FO-AC-103, así que no se revisan ni generan errores. "
                   "Si alguna es en realidad un paz y salvo, revísala aquí.")
        st.dataframe(pd.DataFrame(otros), use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ filtros
conteo = {k: 0 for k in TIPOS}
for d in docs:
    for t in v.categorias_doc(d):
        conteo[t] += 1
conteo["falta_pdf"] = len(faltantes)
disponibles = [k for k, n in conteo.items() if n]
if disponibles:
    st.markdown("#### 🎯 ¿Qué tipo de errores hay?")
    st.markdown("".join(f'<span class="badge" style="background:{TIPOS[k][2]}1F;color:{TIPOS[k][2]};border:1px solid {TIPOS[k][2]}55">'
                        f'{TIPOS[k][0]} {TIPOS[k][1]}: {conteo[k]}</span>' for k in disponibles), unsafe_allow_html=True)
    st.caption("Cada tarjeta compara el FO-AC-83 contra el certificado. Las filas en rojo no coinciden; en amarillo, a revisar.")
    f1, f2, f3 = st.columns([2, 1, 1])
    elegidos = f1.multiselect("Tipo de error", disponibles, default=disponibles,
                              format_func=lambda k: f"{TIPOS[k][0]} {TIPOS[k][1]} ({conteo[k]})")
    busca = f2.text_input("Buscar por nombre o cédula")
    ver_ok = f3.checkbox("Incluir los que coinciden", value=False)

    def pasa(d):
        if busca.strip():
            q = v.normaliza(busca)
            qc = v.solo_digitos(busca)
            if not (q in v.normaliza(d.nombre_firma or d.nombre) or (qc and qc in d.cedula)):
                return False
        tipos = v.categorias_doc(d)
        return (not tipos and ver_ok) or any(t in elegidos for t in tipos)

    vista = [d for d in docs if pasa(d)]
    st.markdown(f"**Mostrando {len(vista)} certificado(s)**")
    for d in vista[:80]:
        st.markdown(card(d), unsafe_allow_html=True)
        if st.toggle("Ver página del PDF", key=f"img{d.archivo}{d.pagina}{d.cedula}"):
            try:
                st.image(v.Renderizador(st.session_state["pdf_bytes"][d.archivo])
                         .render(int(d.pagina.split(",")[0]) - 1, 110))
            except Exception as e:  # noqa: BLE001
                st.error(f"No se pudo mostrar la página: {e}")
    if "falta_pdf" in elegidos and not busca.strip():
        st.markdown(f"#### ⚠️ Retiros del FO-AC-83 sin soporte ({len(faltantes)})")
        for x in faltantes[:80]:
            st.markdown(card_falta_pdf(x), unsafe_allow_html=True)

with st.expander("📋 Ver tablas completas"):
    st.markdown("**Comparación campo por campo**")
    st.dataframe(v.tabla_comparacion(docs), use_container_width=True, hide_index=True)
    st.markdown("**Resumen por certificado**")
    st.dataframe(v.tabla_resumen(docs), use_container_width=True, hide_index=True)

st.download_button("⬇️ Descargar informe (Excel)", v.exportar_excel(res), "revision_paz_y_salvos.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
