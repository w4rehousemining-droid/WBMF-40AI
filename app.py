import os
import tempfile
from io import BytesIO

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
    page_title="WBMF-40AI Manifest & Waybill Generator",
    page_icon="📦",
    layout="wide"
)

TEMPLATE_FILE = "WBMF PO 1011953241 HAJU.xlsx"


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def write_cell(ws, cell_address, value):
    """
    Tulis value ke cell biasa.
    Dipakai untuk Sheet1.
    """
    ws[cell_address] = value if value not in [None, ""] else "-"


def set_cell_safe(ws, cell_address, value):
    """
    Tulis value ke cell tanpa merusak merged cell.
    Jika cell target adalah merged cell non-utama,
    maka value akan ditulis ke cell utama/top-left dari merged range.
    """
    cell = ws[cell_address]

    if not isinstance(cell, MergedCell):
        cell.value = value
        return

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            top_left_cell = ws.cell(
                row=merged_range.min_row,
                column=merged_range.min_col
            )
            top_left_cell.value = value
            return


def clear_cell_safe(ws, cell_address):
    """
    Clear cell tanpa merusak merged cell.
    Jika cell adalah merged cell non-utama, dilewati.
    """
    cell = ws[cell_address]

    if isinstance(cell, MergedCell):
        return

    cell.value = None


def clear_range_safe(ws, min_row, max_row, columns):
    """
    Clear area input tanpa unmerge dan tanpa mengubah layout.
    """
    for row in range(min_row, max_row + 1):
        for col in columns:
            clear_cell_safe(ws, f"{col}{row}")


def clear_picture_sheet(ws_picture, keep_title=True):
    """
    Bersihkan sheet PICTURE dari tulisan/data yang tidak perlu.
    Judul PICTURE & ATTACHMENT di row 1 tetap dipertahankan jika keep_title=True.
    Gambar lama/contoh juga dihapus.
    """

    # Hapus semua gambar lama/contoh
    ws_picture._images = []

    # Simpan title jika ada
    title_value = ws_picture["A1"].value if keep_title else None

    # Bersihkan semua isi cell, kecuali row 1 jika keep_title=True
    for row in ws_picture.iter_rows():
        for cell in row:
            if keep_title and cell.row == 1:
                continue

            if isinstance(cell, MergedCell):
                continue

            cell.value = None

    # Pastikan title tetap ada
    if keep_title:
        ws_picture["A1"] = title_value if title_value else "PICTURE & ATTACHMENT"


def find_row_by_text(ws, text):
    """
    Cari nomor row berdasarkan text tertentu.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == text:
                return cell.row
    return None


def prepare_valid_items(items_df):
    """
    Ambil data material yang valid dari dataframe Streamlit.
    """
    valid_items = []

    for _, row in items_df.iterrows():
        material_desc = str(row.get("Material Description", "")).strip()

        if material_desc == "" or material_desc.lower() == "nan":
            continue

        quantity = row.get("Quantity", 0)
        quantity_received = row.get("Quantity Received", quantity)

        try:
            quantity = float(quantity)
            if quantity.is_integer():
                quantity = int(quantity)
        except Exception:
            quantity = 0

        try:
            quantity_received = float(quantity_received)
            if quantity_received.is_integer():
                quantity_received = int(quantity_received)
        except Exception:
            quantity_received = quantity

        job_site = str(row.get("Job Site", "")).strip()
        destination = str(row.get("Destination", "")).strip()
        uom = str(row.get("UOM", "EA")).strip()

        if job_site == "" or job_site.lower() == "nan":
            job_site = "-"

        if destination == "" or destination.lower() == "nan":
            destination = "-"

        if uom == "" or uom.lower() == "nan":
            uom = "EA"

        valid_items.append({
            "material_desc": material_desc,
            "quantity": quantity,
            "quantity_received": quantity_received,
            "job_site": job_site,
            "destination": destination,
            "uom": uom
        })

    return valid_items


def save_uploaded_image_to_temp(uploaded_file):
    """
    Simpan gambar upload ke temporary file tanpa mengubah resolusi asli.
    Yang disesuaikan hanya ukuran tampilan gambar di Excel.
    """
    uploaded_file.seek(0)

    file_ext = os.path.splitext(uploaded_file.name)[1].lower()

    if file_ext not in [".png", ".jpg", ".jpeg"]:
        file_ext = ".png"

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=file_ext)
    temp_file.write(uploaded_file.getbuffer())
    temp_file.close()

    return temp_file.name


def get_fitted_image_size(image_path, max_width=360, max_height=260):
    """
    Hitung ukuran tampilan gambar agar pas di area Excel.
    Resolusi file asli tidak diubah.
    """
    with Image.open(image_path) as img:
        original_width, original_height = img.size

    if original_width == 0 or original_height == 0:
        return max_width, max_height

    ratio = min(
        max_width / original_width,
        max_height / original_height
    )

    display_width = int(original_width * ratio)
    display_height = int(original_height * ratio)

    return display_width, display_height


def safe_filename(text):
    """
    Bersihkan nama file dari karakter yang tidak aman untuk Windows.
    """
    text = str(text)
    invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']

    for char in invalid_chars:
        text = text.replace(char, "-")

    return text


# =========================================================
# GENERATE EXCEL SESUAI TEMPLATE ASLI
# =========================================================

def generate_excel(
    form_data,
    items_df,
    uploaded_images,
    clear_picture_content=True
):
    """
    Generate Excel dengan mempertahankan format template asli.
    Tidak unmerge.
    Tidak mengubah layout Manifest dan Waybill.
    Sheet PICTURE dibersihkan dari tulisan tidak perlu.
    """

    wb = load_workbook(TEMPLATE_FILE)

    ws_input = wb["Sheet1"]
    ws_manifest = wb["Manifest"]
    ws_waybill = wb["Waybill"]
    ws_picture = wb["PICTURE"]

    # =====================================================
    # UPDATE SHEET1
    # =====================================================

    write_cell(ws_input, "B1", form_data["mf_number"])
    write_cell(ws_input, "B2", form_data["wb_number"])
    write_cell(ws_input, "B3", form_data["po_sto"])
    write_cell(ws_input, "B4", form_data["forwarder"])
    write_cell(ws_input, "B5", form_data["delivery_mode"])
    write_cell(ws_input, "B6", form_data["insurance_po_number"])
    write_cell(ws_input, "B7", form_data["transportation_type"])
    write_cell(ws_input, "B8", form_data["transportation_name"])
    write_cell(ws_input, "B9", form_data["transportation_capacity"])
    write_cell(ws_input, "B10", form_data["etd"])
    write_cell(ws_input, "B11", form_data["eta"])
    write_cell(ws_input, "B12", form_data["actual_arrival_date"])
    write_cell(ws_input, "B13", form_data["from_location"])
    write_cell(ws_input, "B14", form_data["to_location"])
    write_cell(ws_input, "B15", form_data["attention"])
    write_cell(ws_input, "B16", form_data["phone"])
    write_cell(ws_input, "B17", form_data["company"])

    valid_items = prepare_valid_items(items_df)

    # =====================================================
    # UPDATE MANIFEST
    # =====================================================
    # Manifest:
    # A = No.
    # B = Waybill No.
    # C = Description
    # I = Quantity
    # K = Destination
    # P = UOM

    manifest_start_row = 17
    total_qty_row = find_row_by_text(ws_manifest, "Total Quantity")
    manifest_end_row = total_qty_row - 1 if total_qty_row else 130

    clear_range_safe(
        ws_manifest,
        manifest_start_row,
        manifest_end_row,
        ["A", "B", "C", "I", "K", "P"]
    )

    manifest_row = manifest_start_row
    total_quantity = 0

    for idx, item in enumerate(valid_items, start=1):
        if manifest_row > manifest_end_row:
            break

        set_cell_safe(ws_manifest, f"A{manifest_row}", idx)
        set_cell_safe(ws_manifest, f"B{manifest_row}", form_data["wb_number"])
        set_cell_safe(ws_manifest, f"C{manifest_row}", item["material_desc"])
        set_cell_safe(ws_manifest, f"I{manifest_row}", item["quantity"])
        set_cell_safe(ws_manifest, f"K{manifest_row}", item["destination"])
        set_cell_safe(ws_manifest, f"P{manifest_row}", item["uom"])

        total_quantity += item["quantity"]

        # Template Manifest menggunakan jarak 2 row antar item.
        manifest_row += 2

    if total_qty_row:
        set_cell_safe(ws_manifest, f"I{total_qty_row}", total_quantity)

    # =====================================================
    # UPDATE WAYBILL
    # =====================================================
    # Waybill:
    # A = No. Item
    # C = Item Name / Description
    # H = Job Site
    # K = Destination
    # M = Quantity Delivered
    # O = Quantity Received

    waybill_start_row = 8
    prepared_row = find_row_by_text(ws_waybill, "Prepared By")
    waybill_end_row = prepared_row - 1 if prepared_row else 30

    clear_range_safe(
        ws_waybill,
        waybill_start_row,
        waybill_end_row,
        ["A", "C", "H", "K", "M", "O"]
    )

    waybill_row = waybill_start_row

    for idx, item in enumerate(valid_items, start=1):
        if waybill_row > waybill_end_row:
            break

        set_cell_safe(ws_waybill, f"A{waybill_row}", idx)
        set_cell_safe(ws_waybill, f"C{waybill_row}", item["material_desc"])
        set_cell_safe(ws_waybill, f"H{waybill_row}", item["job_site"])
        set_cell_safe(ws_waybill, f"K{waybill_row}", item["destination"])
        set_cell_safe(ws_waybill, f"M{waybill_row}", item["quantity"])
        set_cell_safe(ws_waybill, f"O{waybill_row}", item["quantity_received"])

        waybill_row += 1

    # =====================================================
    # UPDATE PICTURE
    # =====================================================

    if clear_picture_content:
        # Hapus tulisan tidak perlu dan gambar contoh/lama.
        # Judul PICTURE & ATTACHMENT tetap dipertahankan.
        clear_picture_sheet(ws_picture, keep_title=True)

    if uploaded_images:
        picture_slots = [
            {"cell": "B3", "max_width": 360, "max_height": 260},
            {"cell": "K3", "max_width": 360, "max_height": 260},
            {"cell": "B28", "max_width": 360, "max_height": 260},
            {"cell": "K28", "max_width": 360, "max_height": 260},
            {"cell": "B53", "max_width": 360, "max_height": 260},
            {"cell": "K53", "max_width": 360, "max_height": 260},
            {"cell": "B78", "max_width": 360, "max_height": 260},
            {"cell": "K78", "max_width": 360, "max_height": 260},
        ]

        for uploaded_file, slot in zip(uploaded_images, picture_slots):
            temp_image_path = save_uploaded_image_to_temp(uploaded_file)

            img = XLImage(temp_image_path)

            display_width, display_height = get_fitted_image_size(
                temp_image_path,
                max_width=slot["max_width"],
                max_height=slot["max_height"]
            )

            # Ini hanya mengubah ukuran tampilan di Excel,
            # bukan mengubah resolusi asli file gambar.
            img.width = display_width
            img.height = display_height

            ws_picture.add_image(img, slot["cell"])

    # Paksa Excel recalculate formula saat file dibuka.
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, len(valid_items), total_quantity


# =========================================================
# STREAMLIT UI
# =========================================================

st.title("📦 WBMF-40AI Manifest & Waybill Generator")

st.write(
    "Input data dari Streamlit, lalu generate Excel sesuai template asli. "
    "Fitur PDF sudah dihapus."
)

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"Template Excel tidak ditemukan: {TEMPLATE_FILE}")
    st.info("Pastikan file template Excel ada satu folder dengan app.py.")
    st.stop()


with st.sidebar:
    st.header("📁 Setup")
    st.success("Template Excel ditemukan.")
    st.caption(TEMPLATE_FILE)

    st.divider()

    st.header("⚙️ Opsi Picture")
    clear_picture_content = st.checkbox(
        "Bersihkan tulisan/gambar lama di sheet PICTURE",
        value=True
    )

    st.caption(
        "Jika dicentang, tulisan tidak perlu dan gambar contoh di sheet PICTURE akan dihapus. "
        "Judul PICTURE & ATTACHMENT tetap dipertahankan."
    )


tab1, tab2, tab3 = st.tabs([
    "1️⃣ Header",
    "2️⃣ Material & Quantity",
    "3️⃣ Picture Optional"
])


# =========================================================
# TAB 1 - HEADER
# =========================================================

with tab1:
    st.subheader("Input Header Sesuai Sheet1")

    col1, col2 = st.columns(2)

    with col1:
        mf_number = st.text_input(
            "MF Number",
            value="MF40AI-02052026/MACOMINING/02"
        )

        wb_number = st.text_input(
            "WB Number",
            value="WB40AI-02052026/MACOMINING/02"
        )

        po_sto = st.text_input(
            "PO / STO",
            value="1011953241"
        )

        forwarder = st.text_input(
            "Forwarder / Pengirim",
            value="-"
        )

        delivery_mode = st.selectbox(
            "Delivery Mode",
            ["LAND", "SEA", "AIR"],
            index=0
        )

        insurance_po_number = st.text_input(
            "Insurance PO Number",
            value="-"
        )

        transportation_type = st.text_input(
            "Transportation Type",
            value="-"
        )

        transportation_name = st.text_input(
            "Transportation Name",
            value="-"
        )

        transportation_capacity = st.text_input(
            "Transportation Capacity",
            value="-"
        )

    with col2:
        etd = st.text_input(
            "Estimated Time of Departure",
            value="May, 02 2026"
        )

        eta = st.text_input(
            "Estimated Time of Arrival",
            value="May, 02 2026"
        )

        actual_arrival_date = st.text_input(
            "Actual Arrival Date",
            value="-"
        )

        from_location = st.text_area(
            "From",
            value="PT. SAPTAINDRA SEJATI SITE MACO MINING"
        )

        to_location = st.text_area(
            "To",
            value="PT. SAPTAINDRA SEJATI SITE MACO HAULING"
        )

        attention = st.text_input(
            "Attention / Penerima",
            value="Bapak Fachry"
        )

        phone = st.text_input(
            "Phone",
            value="+62 812-8347-1699"
        )

        company = st.text_input(
            "Company",
            value="PT SAPTAINDRA SEJATI"
        )


# =========================================================
# TAB 2 - MATERIAL
# =========================================================

with tab2:
    st.subheader("Input Material Description dan Quantity")

    st.info(
        "Material Description dan Quantity akan masuk ke Manifest dan Waybill "
        "tanpa mengubah layout template Excel."
    )

    default_data = pd.DataFrame([
        {
            "Material Description": "02091-12030",
            "Quantity": 5,
            "Quantity Received": 5,
            "Job Site": "MACO MINING",
            "Destination": "MACO HAULING",
            "UOM": "EA"
        },
        {
            "Material Description": "02290-12031",
            "Quantity": 1,
            "Quantity Received": 1,
            "Job Site": "MACO MINING",
            "Destination": "MACO HAULING",
            "UOM": "EA"
        }
    ])

    edited_items_df = st.data_editor(
        default_data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Material Description": st.column_config.TextColumn(
                "Material Description",
                help="Isi part number, material code, atau deskripsi material",
                required=True
            ),
            "Quantity": st.column_config.NumberColumn(
                "Quantity",
                min_value=0,
                step=1,
                required=True
            ),
            "Quantity Received": st.column_config.NumberColumn(
                "Quantity Received",
                min_value=0,
                step=1
            ),
            "Job Site": st.column_config.TextColumn("Job Site"),
            "Destination": st.column_config.TextColumn("Destination"),
            "UOM": st.column_config.TextColumn("UOM")
        }
    )

    preview_items = prepare_valid_items(edited_items_df)
    total_preview_qty = sum(item["quantity"] for item in preview_items)

    col_a, col_b = st.columns(2)

    with col_a:
        st.metric("Total Item", len(preview_items))

    with col_b:
        st.metric("Total Quantity", total_preview_qty)


# =========================================================
# TAB 3 - PICTURE
# =========================================================

with tab3:
    st.subheader("Upload Picture / Attachment Optional")

    uploaded_images = st.file_uploader(
        "Upload gambar jika ada",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

    if uploaded_images:
        st.success(f"{len(uploaded_images)} gambar dipilih.")

        preview_cols = st.columns(3)

        for idx, uploaded_image in enumerate(uploaded_images):
            with preview_cols[idx % 3]:
                st.image(
                    uploaded_image,
                    caption=uploaded_image.name,
                    width=250
                )
    else:
        st.info("Picture tidak wajib. Jika tidak ada, langsung generate saja.")


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

generate = st.button(
    "🚀 Generate Excel",
    type="primary",
    use_container_width=True
)

if generate:
    valid_items = prepare_valid_items(edited_items_df)

    if len(valid_items) == 0:
        st.warning("Minimal isi 1 Material Description terlebih dahulu.")
    else:
        try:
            excel_file, item_count, total_quantity = generate_excel(
                form_data=form_data,
                items_df=edited_items_df,
                uploaded_images=uploaded_images,
                clear_picture_content=clear_picture_content
            )

            safe_po_sto = safe_filename(po_sto)
            safe_wb_number = safe_filename(wb_number)

            excel_filename = f"WBMF_{safe_po_sto}_{safe_wb_number}.xlsx"

            st.success(
                f"Excel berhasil dibuat. Total item: {item_count}, "
                f"Total quantity: {total_quantity}"
            )

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_file,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as error:
            st.error("Generate file gagal.")
            st.code(str(error))