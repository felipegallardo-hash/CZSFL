from __future__ import annotations

import io
import math
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "Cotizador_Empaques_2026_VERSION_FINAL.xlsx"

st.set_page_config(
    page_title="Cotizador Sellado de Fluidos",
    page_icon="🧾",
    layout="wide",
)

# ---------- Styling ----------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1250px;}
    .main-title {
        background:#17365D; color:white; padding:18px 22px; border-radius:12px;
        font-size:30px; font-weight:700; margin-bottom:8px;
    }
    .sub-title {color:#666; margin-bottom:20px;}
    .result-box {
        background:#E2F0D9; border:1px solid #B7C9D6; border-radius:10px;
        padding:14px 16px; margin-bottom:10px;
    }
    .cost-box {
        background:#D9EAF7; border:1px solid #B7C9D6; border-radius:10px;
        padding:14px 16px; margin-bottom:10px;
    }
    .small-muted {font-size:0.9rem; color:#6b7280;}
    </style>
    """,
    unsafe_allow_html=True,
)

def clp(value: float | int | None) -> str:
    if value is None:
        return "-"
    return "$ " + f"{round(float(value)):,.0f}".replace(",", ".")

def excel_round_hundred(value: float) -> int:
    """Equivalent to Excel ROUND(value,-2) for positive quote values."""
    d = Decimal(str(value)) / Decimal("100")
    return int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP) * Decimal("100"))

def excel_roundup_integer(value: float) -> int:
    """Equivalent to Excel ROUNDUP(value,0) for positive quote values."""
    return int(math.ceil(value))

def sale_price_from_factor(cost: float, factor: float) -> int:
    """Cotizador original: ROUND(costo * factor, -2)."""
    return excel_round_hundred(cost * factor)

def sale_price_from_margin(cost: float, margin_pct: float) -> int:
    """Cotizador original: ROUNDUP(costo / (1 - margen), 0)."""
    margin = margin_pct / 100.0
    if margin >= 1:
        raise ValueError("El margen debe ser menor a 100%.")
    return excel_roundup_integer(cost / (1 - margin))

@st.cache_data(show_spinner=False)
def load_master_data(path: str):
    wb_values = load_workbook(path, data_only=True, read_only=True)

    # ---- Flat gaskets / cuts ----
    listado = wb_values["LISTADO"]

    cuts = {}
    for row in listado.iter_rows(min_row=2, min_col=1, max_col=12, values_only=True):
        key = row[0]       # CONCATENADO, e.g. FF150 3-1/2"
        if not key:
            continue
        cuts[str(key)] = {
            "od_mm": float(row[3]) if row[3] is not None else None,
            "id_mm": float(row[4]) if row[4] is not None else None,
            "class": row[7],
            "cut_cost": float(row[11]) if row[11] is not None else 0.0,
        }

    materials = {}
    for row in listado.iter_rows(min_row=2, min_col=33, max_col=37, values_only=True):
        name, inventory, unit_cost, length, width = row
        if not name or length is None or width is None:
            continue
        materials[str(name)] = {
            "inventory": float(inventory) if inventory is not None else 0.0,
            "sheet_cost": float(unit_cost) if unit_cost is not None else 0.0,
            "length": float(length),
            "width": float(width),
        }

    # ---- Braided packing ----
    # LISTADO AC=style, AD=section, AF=real product code.
    braid_map = []
    for r in range(14, 30):
        style = listado.cell(r, 29).value
        section = listado.cell(r, 30).value
        code = listado.cell(r, 32).value
        if style is not None and section is not None and code:
            braid_map.append((style, str(section), str(code)))

    mayor = wb_values["MAYOR AUX G"]
    cost_inventory = {}
    for row in mayor.iter_rows(min_row=3, min_col=1, max_col=11, values_only=True):
        code = row[0]
        if not code:
            continue
        cost_inventory[str(code).strip("'")] = {
            "description": row[1],
            "stock": float(row[5]) if row[5] is not None else 0.0,
            "cost": float(row[10]) if row[10] is not None else 0.0,
        }

    braids = {}
    for style, section, code in braid_map:
        source = cost_inventory.get(code, {})
        braids[(str(style), section)] = {
            "code": code,
            "cost": float(source.get("cost", 0.0)),
            "stock": float(source.get("stock", 0.0)),
            "description": source.get("description", ""),
        }

    return cuts, materials, braids

def quote_xlsx(items: list[dict], client: str, quote_no: str, quote_date: date) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "COTIZACION"

    navy = "17365D"
    green = "E2F0D9"
    white = "FFFFFF"
    thin = Side(style="thin", color="B7C9D6")

    ws.merge_cells("A1:F1")
    ws["A1"] = "COTIZACIÓN DE EMPAQUETADURAS"
    ws["A1"].font = Font(size=16, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws["A3"], ws["B3"] = "Cliente", client
    ws["D3"], ws["E3"] = "N° Cotización", quote_no
    ws["A4"], ws["B4"] = "Fecha", quote_date.strftime("%d-%m-%Y")

    headers = ["N°", "Tipo", "Glosa", "Cantidad", "Precio unit.", "Subtotal"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(6, c, h)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, item in enumerate(items, 1):
        r = 6 + i
        values = [
            i,
            item["Tipo"],
            item["Glosa"],
            item["Cantidad"],
            item["Precio unitario"],
            item["Subtotal"],
        ]
        for c, v in enumerate(values, 1):
            cell = ws.cell(r, c, v)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(r, 5).number_format = '$ #,##0'
        ws.cell(r, 6).number_format = '$ #,##0'

    total_row = 8 + len(items)
    ws.cell(total_row, 5, "TOTAL")
    ws.cell(total_row, 5).font = Font(bold=True, color=white)
    ws.cell(total_row, 5).fill = PatternFill("solid", fgColor=navy)
    ws.cell(total_row, 6, sum(x["Subtotal"] for x in items))
    ws.cell(total_row, 6).font = Font(bold=True)
    ws.cell(total_row, 6).fill = PatternFill("solid", fgColor=green)
    ws.cell(total_row, 6).number_format = '$ #,##0'

    widths = {"A": 7, "B": 18, "C": 55, "D": 14, "E": 18, "F": 18}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

cuts, materials, braids = load_master_data(str(DATA_FILE))

if "quote_items" not in st.session_state:
    st.session_state.quote_items = []

st.markdown('<div class="main-title">COTIZADOR · SELLADO DE FLUIDOS</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Cotizador de empaquetaduras planas y trenzadas basado en la matriz de costos validada.</div>',
    unsafe_allow_html=True,
)
st.info(
    "Importante: usa **Factor multiplicador** cuando trabajes con valores como 2 o 6,5. "
    "Usa **Margen %** cuando quieras ingresar, por ejemplo, 50% de margen. "
    "La app aplica la misma fórmula del cotizador original para cada caso."
)

# Quote metadata
m1, m2, m3 = st.columns([2.2, 1.2, 1])
with m1:
    client = st.text_input("Cliente", placeholder="Ej.: Compañía Minera...")
with m2:
    quote_no = st.text_input("N° Cotización", placeholder="SFL0826-CZ...")
with m3:
    quote_date = st.date_input("Fecha", value=date.today())

tab_flat, tab_braid, tab_quote = st.tabs(
    ["🟦 Empaquetadura plana", "🟩 Empaquetadura trenzada", "🧾 Cotización"]
)

# ========================= FLAT =========================
with tab_flat:
    st.subheader("Empaquetadura plana / cortada")
    c1, c2 = st.columns(2)

    with c1:
        material = st.selectbox("Material", sorted(materials.keys()))
        cut = st.selectbox("Tipo de corte", list(cuts.keys()))
        qty = st.number_input("Cantidad a cotizar", min_value=1, value=1, step=1)

    with c2:
        pricing_mode_flat = st.radio(
            "Método de precio",
            ["Factor multiplicador", "Margen %"],
            horizontal=True,
            key="pricing_mode_flat",
            help="El cotizador original usa fórmulas distintas para factor y margen.",
        )
        if pricing_mode_flat == "Factor multiplicador":
            factor = st.number_input(
                "Factor",
                min_value=0.01,
                value=2.0,
                step=0.1,
                format="%.2f",
                key="flat_factor",
                help="Ejemplo: 2,0 duplica el costo.",
            )
            margin_pct_flat = None
        else:
            margin_pct_flat = st.number_input(
                "Margen (%)",
                min_value=0.0,
                max_value=99.9,
                value=50.0,
                step=1.0,
                format="%.1f",
                key="flat_margin",
                help="Ejemplo: para 50% ingresa 50.",
            )
            factor = None

    mat = materials[material]
    cut_data = cuts[cut]

    od = cut_data["od_mm"]
    pieces = 0
    if od and od > 0:
        pieces = math.floor(mat["length"] / od) * math.floor(mat["width"] / od)

    cut_cost = cut_data["cut_cost"]
    material_unit_cost = (mat["sheet_cost"] / pieces) if pieces else 0.0
    flat_cost = cut_cost + material_unit_cost
    if pieces:
        if pricing_mode_flat == "Factor multiplicador":
            unit_sale = sale_price_from_factor(flat_cost, factor)
        else:
            unit_sale = sale_price_from_margin(flat_cost, margin_pct_flat)
    else:
        unit_sale = 0
    total = unit_sale * qty
    glosa = f"EMPAQUETADURA {cut} {material}"

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Cantidad por plancha", f"{pieces}")
    r2.metric("Costo corte unit.", clp(cut_cost))
    r3.metric("Costo material / un.", clp(material_unit_cost))
    r4.metric("Inventario planchas", f"{mat['inventory']:,.0f}".replace(",", "."))

    p1, p2 = st.columns(2)
    with p1:
        st.markdown(
            f'<div class="result-box"><b>Precio venta unitario</b><br><span style="font-size:1.65rem">{clp(unit_sale)}</span></div>',
            unsafe_allow_html=True,
        )
    with p2:
        st.markdown(
            f'<div class="result-box"><b>Total cotización</b><br><span style="font-size:1.65rem">{clp(total)}</span></div>',
            unsafe_allow_html=True,
        )

    st.caption(glosa)

    if st.button("➕ Agregar plana a cotización", type="primary", use_container_width=True):
        st.session_state.quote_items.append(
            {
                "Tipo": "PLANA",
                "Glosa": glosa,
                "Cantidad": qty,
                "Unidad": "UN",
                "Precio unitario": unit_sale,
                "Subtotal": total,
            }
        )
        st.success("Producto agregado a la cotización.")

# ========================= BRAIDED =========================
with tab_braid:
    st.subheader("Empaquetadura trenzada")

    styles = list(dict.fromkeys(key[0] for key in braids.keys()))
    b1, b2 = st.columns(2)

    with b1:
        style = st.selectbox("Estilo", styles)
        available_sections = [key[1] for key in braids.keys() if key[0] == style]
        section = st.selectbox("Sección", available_sections)
        kg = st.number_input("Kg a cotizar", min_value=0.01, value=2.0, step=0.5, format="%.2f")

    with b2:
        pricing_mode_braid = st.radio(
            "Método de precio",
            ["Factor multiplicador", "Margen %"],
            horizontal=True,
            key="pricing_mode_braid",
            help="El cotizador original usa fórmulas distintas para factor y margen.",
        )
        if pricing_mode_braid == "Factor multiplicador":
            braid_factor = st.number_input(
                "Factor",
                min_value=0.01,
                value=6.5,
                step=0.1,
                format="%.2f",
                key="braid_factor",
                help="Ejemplo: 6,5 multiplica el costo por kg por 6,5.",
            )
            braid_margin_pct = None
        else:
            braid_margin_pct = st.number_input(
                "Margen (%)",
                min_value=0.0,
                max_value=99.9,
                value=50.0,
                step=1.0,
                format="%.1f",
                key="braid_margin",
                help="Ejemplo: para 50% ingresa 50.",
            )
            braid_factor = None

    braid = braids[(style, section)]
    cost_kg = braid["cost"]
    if pricing_mode_braid == "Factor multiplicador":
        price_kg = sale_price_from_factor(cost_kg, braid_factor)
    else:
        price_kg = sale_price_from_margin(cost_kg, braid_margin_pct)
    braid_total = price_kg * kg
    braid_glosa = f'EMPAQUETADURA TRENZADA ESTILO {style} SECCION {section}'
    display_section = section.replace('"', '')

    a1, a2, a3 = st.columns(3)
    a1.metric("Costo por kg", clp(cost_kg))
    a2.metric("Stock registrado (kg)", f"{braid['stock']:,.2f}".replace(",", "."))
    a3.metric("Código", braid["code"])

    p1, p2 = st.columns(2)
    with p1:
        st.markdown(
            f'<div class="result-box"><b>Precio venta / kg</b><br><span style="font-size:1.65rem">{clp(price_kg)}</span></div>',
            unsafe_allow_html=True,
        )
    with p2:
        st.markdown(
            f'<div class="result-box"><b>Total cotización</b><br><span style="font-size:1.65rem">{clp(braid_total)}</span></div>',
            unsafe_allow_html=True,
        )

    st.caption(f"Estilo {style} · Sección {display_section} · {braid['code']}")

    if st.button("➕ Agregar trenzada a cotización", type="primary", use_container_width=True):
        st.session_state.quote_items.append(
            {
                "Tipo": "TRENZADA",
                "Glosa": braid_glosa,
                "Cantidad": kg,
                "Unidad": "KG",
                "Precio unitario": price_kg,
                "Subtotal": braid_total,
            }
        )
        st.success("Producto agregado a la cotización.")

# ========================= QUOTE =========================
with tab_quote:
    st.subheader("Resumen de cotización")

    if not st.session_state.quote_items:
        st.info("Todavía no has agregado productos.")
    else:
        df = pd.DataFrame(st.session_state.quote_items)
        display_df = df[["Tipo", "Glosa", "Cantidad", "Unidad", "Precio unitario", "Subtotal"]].copy()
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Precio unitario": st.column_config.NumberColumn(format="$ %d"),
                "Subtotal": st.column_config.NumberColumn(format="$ %d"),
            },
        )

        quote_total = int(df["Subtotal"].sum())
        st.markdown(
            f'<div class="result-box"><b>TOTAL GENERAL</b><br><span style="font-size:1.9rem">{clp(quote_total)}</span></div>',
            unsafe_allow_html=True,
        )

        c_download, c_clear = st.columns(2)
        with c_download:
            xlsx_bytes = quote_xlsx(
                st.session_state.quote_items,
                client=client,
                quote_no=quote_no,
                quote_date=quote_date,
            )
            filename = f"Cotizacion_{quote_no or 'Sellado_Fluidos'}.xlsx".replace("/", "-")
            st.download_button(
                "⬇️ Descargar cotización en Excel",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        with c_clear:
            if st.button("🗑️ Vaciar cotización", use_container_width=True):
                st.session_state.quote_items = []
                st.rerun()

st.divider()
st.markdown(
    '<div class="small-muted">Base de costos: Cotizador_Empaques_2026_VERSION_FINAL.xlsx. '
    'Para actualizar los costos de la app, reemplaza ese archivo por una versión actualizada manteniendo el mismo nombre y estructura.</div>',
    unsafe_allow_html=True,
)
