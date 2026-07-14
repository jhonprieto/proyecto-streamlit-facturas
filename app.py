from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

EXCEL_PATH = Path(__file__).parent / "Productos.xlsx"
SHEET_NAME = "Productos"
IGV_RATE = 0.18


def validate_excel_file() -> None:
    if not EXCEL_PATH.exists():
        st.error(f"No se encontro el archivo {EXCEL_PATH.name} en la carpeta del proyecto.")
        st.stop()


@st.cache_data(show_spinner=False)
def load_products() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    for col in ["Precio", "Precio_sinIGV", "stock_Ideal"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "Producto_Web" in df.columns:
        df["Producto_Web"] = df["Producto_Web"].astype(str).str.strip()

    return df


def get_headers() -> list[str]:
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    headers = [cell.value for cell in ws[1]]
    wb.close()
    return [str(h) for h in headers if h is not None]


def get_next_id(df: pd.DataFrame) -> int:
    if "ID" not in df.columns:
        return 1
    ids = pd.to_numeric(df["ID"], errors="coerce").dropna()
    if ids.empty:
        return 1
    return int(ids.max()) + 1


def append_product_to_excel(row_map: dict[str, object]) -> None:
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    headers = [cell.value for cell in ws[1]]

    new_row = [row_map.get(str(header), "") for header in headers]
    ws.append(new_row)
    wb.save(EXCEL_PATH)
    wb.close()



def make_invoice_dataframe(items: list[dict[str, object]]) -> pd.DataFrame:
    rows = []
    for item in items:
        cantidad = int(item["Cantidad"])
        precio = float(item["Precio Unitario"])
        subtotal = cantidad * precio
        rows.append(
            {
                "Producto": item["Producto"],
                "Cantidad": cantidad,
                "Precio Unitario": round(precio, 2),
                "Subtotal": round(subtotal, 2),
            }
        )
    return pd.DataFrame(rows)



def to_invoice_excel_bytes(invoice_df: pd.DataFrame, customer: str, invoice_number: str, issue_date: date) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        invoice_df.to_excel(writer, index=False, sheet_name="Factura")

        workbook = writer.book
        worksheet = writer.sheets["Factura"]

        money_fmt = workbook.add_format({"num_format": "S/ #,##0.00"})
        bold_fmt = workbook.add_format({"bold": True})

        last_row = len(invoice_df) + 2
        subtotal = float(invoice_df["Subtotal"].sum())
        igv = subtotal * IGV_RATE
        total = subtotal + igv

        worksheet.write("F1", "Factura Nro", bold_fmt)
        worksheet.write("G1", invoice_number)
        worksheet.write("F2", "Cliente", bold_fmt)
        worksheet.write("G2", customer)
        worksheet.write("F3", "Fecha", bold_fmt)
        worksheet.write("G3", issue_date.strftime("%d/%m/%Y"))

        worksheet.write(last_row, 2, "Subtotal", bold_fmt)
        worksheet.write_number(last_row, 3, subtotal, money_fmt)
        worksheet.write(last_row + 1, 2, "IGV (18%)", bold_fmt)
        worksheet.write_number(last_row + 1, 3, igv, money_fmt)
        worksheet.write(last_row + 2, 2, "Total", bold_fmt)
        worksheet.write_number(last_row + 2, 3, total, money_fmt)

        worksheet.set_column("A:A", 30)
        worksheet.set_column("B:D", 16)
        worksheet.set_column("F:G", 18)

    return output.getvalue()



def render_invoice_tab(products_df: pd.DataFrame) -> None:
    st.subheader("Generador de facturas")

    if products_df.empty:
        st.warning("No hay productos cargados en el Excel.")
        return

    if "invoice_items" not in st.session_state:
        st.session_state.invoice_items = []

    c1, c2, c3 = st.columns(3)
    with c1:
        customer = st.text_input("Cliente", placeholder="Nombre o razon social")
    with c2:
        invoice_number = st.text_input(
            "Numero de factura",
            value=f"F-{datetime.now().strftime('%Y%m%d-%H%M')}",
        )
    with c3:
        issue_date = st.date_input("Fecha", value=date.today())

    product_names = sorted(products_df["Producto_Web"].dropna().astype(str).unique().tolist())

    with st.form("add_item_form", clear_on_submit=True):
        f1, f2 = st.columns([3, 1])
        with f1:
            selected_product = st.selectbox("Producto", options=product_names)
        with f2:
            quantity = st.number_input("Cantidad", min_value=1, value=1, step=1)

        add_item = st.form_submit_button("Agregar a la factura", use_container_width=True)

    if add_item:
        product_row = products_df.loc[products_df["Producto_Web"] == selected_product].iloc[0]
        unit_price = float(product_row.get("Precio", 0) or 0)

        st.session_state.invoice_items.append(
            {
                "Producto": selected_product,
                "Cantidad": int(quantity),
                "Precio Unitario": unit_price,
            }
        )
        st.success(f"Producto agregado: {selected_product}")

    if st.button("Limpiar factura", type="secondary"):
        st.session_state.invoice_items = []
        st.rerun()

    if not st.session_state.invoice_items:
        st.info("Agrega productos para construir la factura.")
        return

    invoice_df = make_invoice_dataframe(st.session_state.invoice_items)
    subtotal = float(invoice_df["Subtotal"].sum())
    igv = subtotal * IGV_RATE
    total = subtotal + igv

    st.dataframe(invoice_df, use_container_width=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Subtotal", f"S/ {subtotal:,.2f}")
    m2.metric("IGV (18%)", f"S/ {igv:,.2f}")
    m3.metric("Total", f"S/ {total:,.2f}")

    csv_bytes = invoice_df.to_csv(index=False).encode("utf-8")
    xlsx_bytes = to_invoice_excel_bytes(invoice_df, customer, invoice_number, issue_date)

    d1, d2 = st.columns(2)
    d1.download_button(
        "Descargar factura (CSV)",
        data=csv_bytes,
        file_name=f"factura_{invoice_number}.csv",
        mime="text/csv",
        use_container_width=True,
    )
    d2.download_button(
        "Descargar factura (Excel)",
        data=xlsx_bytes,
        file_name=f"factura_{invoice_number}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )



def render_product_tab(products_df: pd.DataFrame) -> None:
    st.subheader("Agregar productos al Excel")
    st.caption("Los productos nuevos se guardan automaticamente en Productos.xlsx")

    headers = get_headers()
    next_id = get_next_id(products_df)

    with st.form("new_product_form", clear_on_submit=True):
        c1, c2 = st.columns(2)

        with c1:
            product_name = st.text_input("Nombre del producto", placeholder="Ej: Silla ergonomica")
            color = st.text_input("Color", placeholder="Ej: Negro")
            kardex = st.text_input("Kardex", placeholder="Codigo interno")
            gama = st.text_input("Gama", placeholder="Ej: Premium")

        with c2:
            price = st.number_input("Precio (con IGV)", min_value=0.0, value=0.0, step=1.0)
            stock_ideal = st.number_input("Stock ideal", min_value=0, value=0, step=1)
            karde_eq = st.text_input("Karde_EQ", placeholder="Codigo relacionado")

        submit = st.form_submit_button("Guardar producto", use_container_width=True)

    if submit:
        name_normalized = product_name.strip()
        if not name_normalized:
            st.error("Debes ingresar el nombre del producto.")
            return

        if "Producto_Web" in products_df.columns:
            duplicate = products_df["Producto_Web"].astype(str).str.strip().str.lower() == name_normalized.lower()
            if duplicate.any():
                st.error("Ese producto ya existe en el Excel.")
                return

        row_map = {
            "ID": next_id,
            "Nombre_Drive": name_normalized,
            "Producto_Web": name_normalized,
            "Color": color.strip(),
            "Kardex": kardex.strip(),
            "Nombre_Genesys": name_normalized,
            "Gama": gama.strip(),
            "Precio": float(price),
            "stock_Ideal": int(stock_ideal),
            "Precio_sinIGV": round(float(price) / (1 + IGV_RATE), 2) if price else 0,
            "Karde_EQ": karde_eq.strip(),
        }

        safe_row_map = {h: row_map.get(h, "") for h in headers}
        append_product_to_excel(safe_row_map)
        load_products.clear()
        st.success("Producto guardado correctamente en Productos.xlsx")
        st.rerun()

    st.markdown("### Productos actuales")
    st.dataframe(products_df, use_container_width=True)



def main() -> None:
    st.set_page_config(page_title="Facturador Streamlit", page_icon="🧾", layout="wide")
    st.title("App de Facturacion")
    st.caption("Genera facturas desde tu Excel de productos y agrega nuevos productos al mismo archivo.")

    validate_excel_file()
    products_df = load_products()

    tab1, tab2 = st.tabs(["Generar factura", "Gestion de productos"])

    with tab1:
        render_invoice_tab(products_df)

    with tab2:
        render_product_tab(products_df)


if __name__ == "__main__":
    main()
