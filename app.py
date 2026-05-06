import os
import tempfile
from io import BytesIO
from copy import copy

import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.cell.cell import MergedCell


# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="WBMF-40AI Logistics Robot",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="collapsed"
)

TEMPLATE_FILE = "WBMF PO 1011953241 HAJU.xlsx"


# =========================================================
# STYLE
# =========================================================

st.markdown(
    """
    <style>
    [data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(0,255,170,0.13), transparent 28%),
            radial-gradient(circle at top right, rgba(0,162,255,0.13), transparent 30%),
            linear-gradient(135deg, #071018 0%, #0b1220 50%, #111827 100%);
        color: #e5f7ff;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.4rem;
        padding-bottom: 2rem;
    }

    .robot-header {
        border: 1px solid rgba(0,255,170,0.35);
        background: linear-gradient(135deg, rgba(13,27,42,0.94), rgba(17,24,39,0.94));
        border-radius: 22px;
        padding: 24px 28px;
        box-shadow: 0 0 28px rgba(0,255,170,0.12);
        margin-bottom: 20px;
    }

    .robot-title {
        font-size: 36px;
        font-weight: 900;
        color: #dffcff;
        margin: 0;
    }

    .robot-subtitle {
        color: #92f7d5;
        font-size: 15px;
        margin-top: 8px;
    }

    .robot-badge {
        display: inline-block;
        padding: 7px 12px;
        border-radius: 999px;
        background: rgba(0,255,170,0.10);
        border: 1px solid rgba(0,255,170,0.35);
        color: #7cffd2;
        font-size: 13px;
        margin-right: 8px;
        margin-top: 14px;
    }

    .stTabs [data-baseweb="tab"] {
        background: rgba(15,23,42,0.88);
        border: 1px solid rgba(56,189,248,0.25);
        border-radius: 14px;
        color: #c7f9ff;
        padding: 12px 18px;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, rgba(0,255,170,0.20), rgba(56,189,248,0.20)) !important;
        border: 1px solid rgba(0,255,170,0.75) !important;
        color: #ffffff !important;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius: 14px;
        border: 1px solid rgba(0,255,170,0.55);
        background: linear-gradient(135deg, #00b894, #0984e3);
        color: white;
        font-weight: 800;
        padding: 0.75rem 1rem;
    }

    [data-testid="stMetric"] {
        background: rgba(15,23,42,0.88);
        border: 1px solid rgba(0,255,170,0.25);
        border-radius: 16px;
        padding: 16px;
    }

    [data-testid="stFileUploader"] {
        background: rgba(15,23,42,0.82);
        border: 1px dashed rgba(0,255,170,0.55);
        border-radius: 20px;
        padding: 20px;
    }

    textarea {
        font-family: Consolas, monospace !important;
    }

    h1, h2, h3, label, p, span {
        color: #d9f8ff;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# BASIC HELPERS
# =========================================================

def set_cell_safe(ws, address, value):
    """
    Isi cell biasa atau merged cell.
    Jika target address adalah bagian merged cell, value ditulis ke top-left cell.
    """
    value = value if value not in [None, ""] else "-"
    cell = ws[address]

    if not isinstance(cell, MergedCell):
        cell.value = value
        return

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            ws.cell(
                row=merged_range.min_row,
                column=merged_range.min_col
            ).value = value
            return


def find_row_by_text(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text:
                return cell.row
    return None


def safe_filename(text):
    text = str(text)

    for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        text = text.replace(char, "-")

    return text


def split_lines(text):
    return [
        line.strip()
        for line in str(text).splitlines()
        if line.strip() != ""
    ]


def get_line(lines, index, default=""):
    if index < len(lines):
        return lines[index].strip()

    return default


def normalize_text(value, default="-"):
    value = str(value).strip()

    if value == "" or value.lower() == "nan":
        return default

    return value


def normalize_qty(value, default=1):
    try:
        value = str(value).strip().replace(",", ".")

        if value == "":
            return default

        number = float(value)

        if number.is_integer():
            return int(number)

        return number

    except Exception:
        return default


# =========================================================
# FORMAT / MERGE HELPERS
# =========================================================

def unhide_rows(ws, start_row, end_row):
    for row in range(start_row, end_row + 1):
        ws.row_dimensions[row].hidden = False


def unmerge_overlapping_ranges(ws, start_row, end_row):
    """
    Unmerge hanya area detail yang bersentuhan dengan range row detail.
    Header tidak diubah.
    """
    merged_ranges = list(ws.merged_cells.ranges)

    for merged_range in merged_ranges:
        if merged_range.max_row >= start_row and merged_range.min_row <= end_row:
            ws.unmerge_cells(str(merged_range))


def merge_if_not_merged(ws, cell_range):
    current_merges = [str(rng) for rng in ws.merged_cells.ranges]

    if cell_range not in current_merges:
        ws.merge_cells(cell_range)


def apply_manifest_detail_merges(ws, row):
    """
    Manifest detail:
    A   = No.
    B   = Waybill No.
    C:H = Description
    I:J = Quantity
    K:O = Destination
    P:Q = UOM
    """
    merge_if_not_merged(ws, f"C{row}:H{row}")
    merge_if_not_merged(ws, f"I{row}:J{row}")
    merge_if_not_merged(ws, f"K{row}:O{row}")
    merge_if_not_merged(ws, f"P{row}:Q{row}")


def apply_waybill_detail_merges(ws, row):
    """
    Waybill detail:
    A:B = No. Item
    C:G = Description
    H:J = Job Site
    K:L = Destination
    M:N = Quantity Delivered
    O:P = Quantity Received
    """
    merge_if_not_merged(ws, f"A{row}:B{row}")
    merge_if_not_merged(ws, f"C{row}:G{row}")
    merge_if_not_merged(ws, f"H{row}:J{row}")
    merge_if_not_merged(ws, f"K{row}:L{row}")
    merge_if_not_merged(ws, f"M{row}:N{row}")
    merge_if_not_merged(ws, f"O{row}:P{row}")


def copy_cell_style(source, target):
    if source.has_style:
        target._style = copy(source._style)

    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format


def copy_row_format(ws, source_row, target_row, max_col=17):
    """
    Copy format row tanpa value.
    """
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    ws.row_dimensions[target_row].hidden = False

    for col in range(1, max_col + 1):
        source = ws.cell(row=source_row, column=col)
        target = ws.cell(row=target_row, column=col)
        copy_cell_style(source, target)


def clear_row_values(ws, row, max_col=17):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)

        if not isinstance(cell, MergedCell):
            cell.value = None


def ensure_detail_rows(ws, start_row, marker_text, item_count, style_row):
    """
    Pastikan row detail cukup.
    1 item = 1 row.
    Jika kurang, insert row sebelum marker row.
    """
    marker_row = find_row_by_text(ws, marker_text)

    if marker_row is None:
        marker_row = start_row + 30

    end_row = marker_row - 1
    required_end_row = start_row + item_count - 1

    if item_count > 0 and required_end_row > end_row:
        rows_to_add = required_end_row - end_row
        ws.insert_rows(marker_row, rows_to_add)

        for row in range(marker_row, marker_row + rows_to_add):
            copy_row_format(ws, style_row, row, max_col=17)

        marker_row = find_row_by_text(ws, marker_text)
        end_row = marker_row - 1

    unhide_rows(ws, start_row, end_row)
    unmerge_overlapping_ranges(ws, start_row, end_row)

    for row in range(start_row, end_row + 1):
        copy_row_format(ws, style_row, row, max_col=17)
        clear_row_values(ws, row, max_col=17)

    return start_row, end_row, marker_row


# =========================================================
# INPUT PARSER
# =========================================================

def parse_items(
    description_text,
    qty_delivered_text,
    qty_received_text,
    job_site_text,
    destination_text,
    uom_text
):
    descriptions = split_lines(description_text)
    qty_delivered_lines = split_lines(qty_delivered_text)
    qty_received_lines = split_lines(qty_received_text)
    job_site_lines = split_lines(job_site_text)
    destination_lines = split_lines(destination_text)
    uom_lines = split_lines(uom_text)

    items = []

    for index, description in enumerate(descriptions):
        description = normalize_text(description, default="")

        if description == "":
            continue

        qty_delivered = normalize_qty(
            get_line(qty_delivered_lines, index, ""),
            default=1
        )

        qty_received = normalize_qty(
            get_line(qty_received_lines, index, ""),
            default=qty_delivered
        )

        job_site = normalize_text(
            get_line(job_site_lines, index, ""),
            default="MACO MINING"
        )

        destination = normalize_text(
            get_line(destination_lines, index, ""),
            default="MACO HAULING"
        )

        uom = normalize_text(
            get_line(uom_lines, index, ""),
            default="EA"
        )

        items.append({
            "description": description,
            "quantity_delivered": qty_delivered,
            "quantity_received": qty_received,
            "job_site": job_site,
            "destination": destination,
            "uom": uom
        })

    return items


def preview_dataframe(items):
    return pd.DataFrame([
        {
            "No": idx,
            "Description / Item Name": item["description"],
            "Job Site": item["job_site"],
            "Destination": item["destination"],
            "Quantity Delivered": item["quantity_delivered"],
            "Quantity Received": item["quantity_received"],
            "UOM": item["uom"]
        }
        for idx, item in enumerate(items, start=1)
    ])


# =========================================================
# PICTURE HELPERS
# =========================================================

def clear_picture_sheet(ws_picture):
    ws_picture._images = []

    title_value = ws_picture["A1"].value or "PICTURE & ATTACHMENT"

    for row in ws_picture.iter_rows():
        for cell in row:
            if cell.row == 1:
                continue

            if isinstance(cell, MergedCell):
                continue

            cell.value = None

    ws_picture["A1"] = title_value


def save_uploaded_image(uploaded_file):
    uploaded_file.seek(0)

    extension = os.path.splitext(uploaded_file.name)[1].lower()

    if extension not in [".png", ".jpg", ".jpeg"]:
        extension = ".png"

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=extension)
    temp_file.write(uploaded_file.getbuffer())
    temp_file.close()

    return temp_file.name


def fit_image_size(image_path, max_width=430, max_height=330):
    with Image.open(image_path) as image:
        width, height = image.size

    ratio = min(max_width / width, max_height / height)

    return int(width * ratio), int(height * ratio)


def insert_pictures(ws_picture, uploaded_images):
    clear_picture_sheet(ws_picture)

    if not uploaded_images:
        return

    slots = [
        "B3", "K3",
        "B31", "K31",
        "B59", "K59",
        "B87", "K87"
    ]

    for uploaded_file, cell in zip(uploaded_images, slots):
        image_path = save_uploaded_image(uploaded_file)

        excel_image = XLImage(image_path)
        width, height = fit_image_size(image_path)

        excel_image.width = width
        excel_image.height = height

        ws_picture.add_image(excel_image, cell)


# =========================================================
# HEADER WRITER
# =========================================================

def write_manifest_header(ws, form):
    set_cell_safe(ws, "E3", form["mf_number"])
    set_cell_safe(ws, "E4", form["po_sto"])
    set_cell_safe(ws, "E5", form["insurance_po_number"])
    set_cell_safe(ws, "E6", form["forwarder"])
    set_cell_safe(ws, "E7", form["delivery_mode"])
    set_cell_safe(ws, "E8", form["transportation_type"])
    set_cell_safe(ws, "E9", form["transportation_name"])
    set_cell_safe(ws, "E10", form["transportation_capacity"])
    set_cell_safe(ws, "E11", form["etd"])
    set_cell_safe(ws, "E12", form["eta"])
    set_cell_safe(ws, "E13", form["actual_arrival_date"])

    set_cell_safe(ws, "L3", form["from_location"])
    set_cell_safe(ws, "L8", form["to_location"])
    set_cell_safe(ws, "L11", form["attention"])
    set_cell_safe(ws, "L12", form["phone"])
    set_cell_safe(ws, "L13", form["company"])


def write_waybill_header(ws, form):
    set_cell_safe(ws, "D3", form["wb_number"])


def remove_sheet1_if_exists(wb):
    if "Sheet1" in wb.sheetnames:
        wb.remove(wb["Sheet1"])

    if "Manifest" in wb.sheetnames:
        wb.active = wb.sheetnames.index("Manifest")


# =========================================================
# GENERATE EXCEL
# =========================================================

def generate_excel(form, items, uploaded_images):
    wb = load_workbook(TEMPLATE_FILE)

    ws_manifest = wb["Manifest"]
    ws_waybill = wb["Waybill"]
    ws_picture = wb["PICTURE"]

    # Header direct tanpa Sheet1
    write_manifest_header(ws_manifest, form)
    write_waybill_header(ws_waybill, form)

    # =====================================================
    # MANIFEST DETAIL
    # 1 item = 1 row, merge hanya kolom yang perlu
    # =====================================================

    manifest_start, manifest_end, total_qty_row = ensure_detail_rows(
        ws=ws_manifest,
        start_row=17,
        marker_text="Total Quantity",
        item_count=len(items),
        style_row=17
    )

    manifest_inserted = 0
    total_quantity = 0

    for idx, item in enumerate(items, start=1):
        row = manifest_start + idx - 1

        if row > manifest_end:
            break

        ws_manifest.row_dimensions[row].hidden = False
        apply_manifest_detail_merges(ws_manifest, row)

        set_cell_safe(ws_manifest, f"A{row}", idx)
        set_cell_safe(ws_manifest, f"B{row}", form["wb_number"])
        set_cell_safe(ws_manifest, f"C{row}", item["description"])
        set_cell_safe(ws_manifest, f"I{row}", item["quantity_delivered"])
        set_cell_safe(ws_manifest, f"K{row}", item["destination"])
        set_cell_safe(ws_manifest, f"P{row}", item["uom"])

        total_quantity += item["quantity_delivered"]
        manifest_inserted += 1

    if total_qty_row:
        set_cell_safe(ws_manifest, f"I{total_qty_row}", total_quantity)

    # =====================================================
    # WAYBILL DETAIL
    # 1 item = 1 row, merge hanya kolom yang perlu
    # =====================================================

    waybill_start, waybill_end, prepared_row = ensure_detail_rows(
        ws=ws_waybill,
        start_row=8,
        marker_text="Prepared By",
        item_count=len(items),
        style_row=8
    )

    waybill_inserted = 0

    for idx, item in enumerate(items, start=1):
        row = waybill_start + idx - 1

        if row > waybill_end:
            break

        ws_waybill.row_dimensions[row].hidden = False
        apply_waybill_detail_merges(ws_waybill, row)

        set_cell_safe(ws_waybill, f"A{row}", idx)
        set_cell_safe(ws_waybill, f"C{row}", item["description"])
        set_cell_safe(ws_waybill, f"H{row}", item["job_site"])
        set_cell_safe(ws_waybill, f"K{row}", item["destination"])
        set_cell_safe(ws_waybill, f"M{row}", item["quantity_delivered"])
        set_cell_safe(ws_waybill, f"O{row}", item["quantity_received"])

        waybill_inserted += 1

    # Picture
    insert_pictures(ws_picture, uploaded_images)

    # Remove Sheet1 from output
    remove_sheet1_if_exists(wb)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return {
        "file": output,
        "total_items": len(items),
        "total_quantity": total_quantity,
        "manifest_inserted": manifest_inserted,
        "waybill_inserted": waybill_inserted
    }


# =========================================================
# UI
# =========================================================

st.markdown(
    """
    <div class="robot-header">
        <p class="robot-title">🤖 WBMF-40AI Logistics Robot</p>
        <div class="robot-subtitle">
            Direct input to Manifest and Waybill. Detail 1 baris per material, merge hanya kolom yang perlu.
        </div>
        <span class="robot-badge">📦 Manifest Row 17,18,19...</span>
        <span class="robot-badge">🚚 Waybill Row 8,9,10...</span>
        <span class="robot-badge">🧾 Sheet1 Removed</span>
        <span class="robot-badge">🖼️ Picture Auto Clean</span>
    </div>
    """,
    unsafe_allow_html=True
)

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"Template Excel tidak ditemukan: {TEMPLATE_FILE}")
    st.info("Pastikan file template Excel berada satu folder dengan app.py.")
    st.stop()


tab1, tab2, tab3 = st.tabs([
    "🤖 Header",
    "📦 Excel Column Input",
    "🖼️ Picture"
])


# =========================================================
# HEADER
# =========================================================

with tab1:
    st.subheader("🤖 Header Manifest dan Waybill")

    col1, col2 = st.columns(2)

    with col1:
        mf_number = st.text_input("MF Number", "MF40AI-02052026/MACOMINING/02")
        wb_number = st.text_input("WB Number", "WB40AI-02052026/MACOMINING/02")
        po_sto = st.text_input("PO / STO", "1011953241")
        forwarder = st.text_input("Forwarder / Pengirim", "-")
        delivery_mode = st.selectbox("Delivery Mode", ["LAND", "SEA", "AIR"])
        insurance_po_number = st.text_input("Insurance PO Number", "-")
        transportation_type = st.text_input("Transportation Type", "-")
        transportation_name = st.text_input("Transportation Name", "-")
        transportation_capacity = st.text_input("Transportation Capacity", "-")

    with col2:
        etd = st.text_input("Estimated Time of Departure", "May, 02 2026")
        eta = st.text_input("Estimated Time of Arrival", "May, 02 2026")
        actual_arrival_date = st.text_input("Actual Arrival Date", "-")
        from_location = st.text_area("From", "PT. SAPTAINDRA SEJATI SITE MACO MINING")
        to_location = st.text_area("To", "PT. SAPTAINDRA SEJATI SITE MACO HAULING")
        attention = st.text_input("Attention / Penerima", "Bapak Fachry")
        phone = st.text_input("Phone", "+62 812-8347-1699")
        company = st.text_input("Company", "PT SAPTAINDRA SEJATI")


# =========================================================
# ITEM INPUT
# =========================================================

with tab2:
    st.subheader("📦 Input Kolom Excel")
    st.caption("Copy kolom dari Excel ke bawah. Output detail 1 item = 1 baris, merge hanya kolom yang perlu.")

    col_desc, col_qty_del, col_qty_rec = st.columns([3, 1, 1])

    with col_desc:
        description_text = st.text_area(
            "Description / Item Name",
            height=300,
            placeholder="Paste kolom Description / Item Name"
        )

    with col_qty_del:
        qty_delivered_text = st.text_area(
            "Quantity Delivered",
            height=300,
            placeholder="Paste Qty Delivered"
        )

    with col_qty_rec:
        qty_received_text = st.text_area(
            "Quantity Received",
            height=300,
            placeholder="Opsional"
        )

    col_job, col_dest, col_uom = st.columns([2, 2, 1])

    with col_job:
        job_site_text = st.text_area(
            "Job Site",
            height=220,
            placeholder="Opsional. Default: MACO MINING"
        )

    with col_dest:
        destination_text = st.text_area(
            "Destination",
            height=220,
            placeholder="Opsional. Default: MACO HAULING"
        )

    with col_uom:
        uom_text = st.text_area(
            "UOM",
            height=220,
            placeholder="Opsional. Default: EA"
        )

    preview_items = parse_items(
        description_text=description_text,
        qty_delivered_text=qty_delivered_text,
        qty_received_text=qty_received_text,
        job_site_text=job_site_text,
        destination_text=destination_text,
        uom_text=uom_text
    )

    preview_df = preview_dataframe(preview_items)
    total_qty = sum(item["quantity_delivered"] for item in preview_items)

    metric_1, metric_2, metric_3 = st.columns(3)

    with metric_1:
        st.metric("Total Item", len(preview_items))

    with metric_2:
        st.metric("Total Quantity", total_qty)

    with metric_3:
        st.metric("Detail Mode", "1 Row + Merge")

    if not preview_df.empty:
        st.dataframe(preview_df, use_container_width=True, hide_index=True)
    else:
        st.info("Paste minimal kolom Description / Item Name terlebih dahulu.")


# =========================================================
# PICTURE
# =========================================================

with tab3:
    st.subheader("🖼️ Picture Attachment")

    uploaded_images = st.file_uploader(
        "Upload gambar attachment",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

    if uploaded_images:
        cols = st.columns(4)

        for idx, uploaded_image in enumerate(uploaded_images):
            with cols[idx % 4]:
                st.image(
                    uploaded_image,
                    caption=uploaded_image.name,
                    use_container_width=True
                )
    else:
        st.info("Picture optional. Sheet PICTURE tetap dibersihkan otomatis.")


# =========================================================
# GENERATE
# =========================================================

st.divider()

form_data = {
    "mf_number": mf_number,
    "wb_number": wb_number,
    "po_sto": po_sto,
    "forwarder": forwarder,
    "delivery_mode": delivery_mode,
    "insurance_po_number": insurance_po_number,
    "transportation_type": transportation_type,
    "transportation_name": transportation_name,
    "transportation_capacity": transportation_capacity,
    "etd": etd,
    "eta": eta,
    "actual_arrival_date": actual_arrival_date,
    "from_location": from_location,
    "to_location": to_location,
    "attention": attention,
    "phone": phone,
    "company": company
}

if st.button("🤖 Generate Excel", type="primary", use_container_width=True):
    items = parse_items(
        description_text=description_text,
        qty_delivered_text=qty_delivered_text,
        qty_received_text=qty_received_text,
        job_site_text=job_site_text,
        destination_text=destination_text,
        uom_text=uom_text
    )

    if not items:
        st.warning("Minimal paste 1 Description / Item Name.")
    else:
        try:
            result = generate_excel(
                form=form_data,
                items=items,
                uploaded_images=uploaded_images
            )

            filename = f"WBMF_{safe_filename(po_sto)}_{safe_filename(wb_number)}.xlsx"

            st.success(
                f"Excel berhasil dibuat. "
                f"Item: {result['total_items']} | "
                f"Manifest: {result['manifest_inserted']} | "
                f"Waybill: {result['waybill_inserted']} | "
                f"Total Qty: {result['total_quantity']}"
            )

            st.download_button(
                label="⬇️ Download Excel",
                data=result["file"],
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as error:
            st.error("Generate gagal.")
            st.code(str(error))