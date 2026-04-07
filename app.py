import streamlit as st
import pandas as pd
import io
import os
import math
import re
from pathlib import Path

st.set_page_config(page_title="SEST", layout="wide")
st.title("SEST")

st.markdown("""
<style>
.stTabs [data-baseweb="tab-list"] {
    display: flex !important;
    flex-direction: row !important;
    overflow-x: auto;
    gap: 6px;
}
.stTabs [data-baseweb="tab"] {
    padding: 6px 10px !important;
    font-size: 14px !important;
    white-space: nowrap;
}
</style>
""", unsafe_allow_html=True)

PROJECTS_DIR = Path("projects")
PROJECTS_DIR.mkdir(exist_ok=True)

SUPPLIERS_DIR = Path("suppliers")
SUPPLIERS_DIR.mkdir(exist_ok=True)

PROFILES_FILE = "Profiles.xlsx"
MAX_PIECE_LENGTH = 23.0
DEFAULT_PROJECT_NAME = "default_project"


def safe_name(name):
    name = str(name).strip()
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", name)
    return name if name else "default"


def safe_project_name(name):
    cleaned = safe_name(name)
    return cleaned if cleaned else DEFAULT_PROJECT_NAME


def get_project_results_file(project_name):
    return PROJECTS_DIR / f"{safe_project_name(project_name)}_results.xlsx"


def get_supplier_file(supplier_name):
    return SUPPLIERS_DIR / f"{safe_name(supplier_name)}.xlsx"


def get_saved_project_names():
    return sorted([
        p.name.replace("_results.xlsx", "")
        for p in PROJECTS_DIR.glob("*_results.xlsx")
    ])


def to_float(value, default=0.0):
    if pd.isna(value):
        return default
    if isinstance(value, str):
        value = value.strip().replace(",", ".")
        if value == "":
            return default
    try:
        return float(value)
    except:
        return default


def split_length_and_quantity(length, number, max_piece_length=23.0):
    length = to_float(length)
    number = to_float(number)

    if length <= 0 or number <= 0:
        return 0.0, 0.0, 1

    if length <= max_piece_length:
        return length, number, 1

    pieces_per_item = math.ceil(length / max_piece_length)
    split_length = length / pieces_per_item
    new_number = number * pieces_per_item

    return split_length, new_number, pieces_per_item


def get_zbsl(profile_row, calc_length):
    if calc_length <= 5:
        return to_float(profile_row.get("Lte5", 0))
    elif calc_length <= 8:
        return to_float(profile_row.get("L5to8", 0))
    elif calc_length <= 11:
        return to_float(profile_row.get("L8to11", 0))
    elif calc_length <= 14:
        return to_float(profile_row.get("L11to14", 0))
    elif calc_length <= 18:
        return to_float(profile_row.get("L14to18", 0))
    else:
        return to_float(profile_row.get("Gt18", 0))


def get_profile_type(profile_name):
    profile_name = str(profile_name).strip().upper()

    if profile_name.startswith(("HEA", "HEB", "HEM", "IPE", "IPN", "INP")):
        return "I Profile"
    elif profile_name.startswith(("K", "RHS", "SHS")):
        return "RHS Profile"
    elif profile_name.startswith("L"):
        return "L Profile"
    elif profile_name.startswith(("UPE", "UNP", "UPN")):
        return "U Profile"
    elif profile_name.startswith(("R", "CHS")):
        return "CHS Profile"
    elif profile_name.startswith("PL"):
        return "PL Profile"
    else:
        return "Other"


def get_weight_factor(profile_name):
    p = str(profile_name).strip().upper()

    if p.startswith(("HEA", "HEB", "HEM", "IPE", "IPN", "UPN", "INP")):
        return 1.15
    elif p.startswith(("K", "L", "R", "RHS", "SHS", "CHS")):
        return 1.20
    elif p.startswith("PL"):
        return 1.40
    else:
        return 1.00


def load_profiles():
    if not os.path.exists(PROFILES_FILE):
        st.error(f"{PROFILES_FILE} not found.")
        st.stop()

    df = pd.read_excel(PROFILES_FILE, header=0)
    df.columns = df.columns.str.strip()

    required_cols = ["Profile", "kgm", "m2_per_m", "Lte5", "L5to8", "L8to11", "L11to14", "L14to18", "Gt18"]
    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        st.error(f"Missing columns in {PROFILES_FILE}: {missing}")
        st.write("Columns found:", df.columns.tolist())
        st.stop()

    return df


def load_supplier_names():
    return sorted([p.stem for p in SUPPLIERS_DIR.glob("*.xlsx")])


def load_supplier_data_by_name(supplier_name):
    if not supplier_name:
        return pd.DataFrame(columns=["Supplier", "Profile Type", "Fabric Standard Length"])

    supplier_file = get_supplier_file(supplier_name)

    if supplier_file.exists():
        df = pd.read_excel(supplier_file).fillna("")
        expected_cols = ["Supplier", "Profile Type", "Fabric Standard Length"]
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ""
        return df[expected_cols]

    return pd.DataFrame(columns=["Supplier", "Profile Type", "Fabric Standard Length"])


def save_supplier_data_by_name(supplier_name, df):
    supplier_file = get_supplier_file(supplier_name)
    df.to_excel(supplier_file, index=False)


def get_supplier_row(profile_type, supplier_df):
    match = supplier_df[
        supplier_df["Profile Type"].astype(str).str.strip() == str(profile_type).strip()
    ]
    if not match.empty:
        return match.iloc[0]
    return None


def rename_project_file(old_name, new_name):
    old_results = get_project_results_file(old_name)
    new_results = get_project_results_file(new_name)
    if old_results.exists():
        old_results.rename(new_results)


def calculate_row(row_data, profile_df):
    profile_name = str(row_data.get("Profile", "")).strip()

    default_result = {
        "Split Pieces": 1,
        "Split Length": 0.0,
        "Split Quantity": 0.0,
        "Total Treatment Area": 0.0,
        "Net Weight": 0.0,
        "Weight Incl. Waste": 0.0,
        "Total ZBSL": 0.0,
        "Total Levering Price": 0.0,
    }

    for k, v in default_result.items():
        row_data[k] = v

    if profile_name == "":
        return row_data

    profile_match = profile_df[profile_df["Profile"].astype(str).str.strip() == profile_name]
    if profile_match.empty:
        return row_data

    profile_row = profile_match.iloc[0]

    input_length = to_float(row_data.get("Length", 0))
    input_quantity = to_float(row_data.get("Quantity", 0))
    price_per_ton = to_float(row_data.get("Price/t", 0))

    calc_length, calc_quantity, split_pieces = split_length_and_quantity(
        input_length, input_quantity, MAX_PIECE_LENGTH
    )

    kgm = to_float(profile_row.get("kgm", 0))
    m2_per_m = to_float(profile_row.get("m2_per_m", 0))
    zbsl = get_zbsl(profile_row, calc_length)

    net_weight = kgm * calc_length * calc_quantity
    factor = get_weight_factor(profile_name)
    weight_incl_waste = net_weight * factor

    total_treatment_area = m2_per_m * calc_length * calc_quantity
    total_zbsl = zbsl * calc_quantity
    total_price = (weight_incl_waste / 1000) * price_per_ton

    row_data["Length"] = round(input_length, 2)
    row_data["Quantity"] = int(input_quantity) if float(input_quantity).is_integer() else round(input_quantity, 2)
    row_data["Split Pieces"] = int(split_pieces)
    row_data["Split Length"] = round(calc_length, 2)
    row_data["Split Quantity"] = int(calc_quantity) if float(calc_quantity).is_integer() else round(calc_quantity, 2)
    row_data["Price/t"] = round(price_per_ton, 2)
    row_data["Total Treatment Area"] = round(total_treatment_area, 2)
    row_data["Net Weight"] = round(net_weight, 2)
    row_data["Weight Incl. Waste"] = round(weight_incl_waste, 2)
    row_data["Total ZBSL"] = round(total_zbsl, 2)
    row_data["Total Levering Price"] = round(total_price, 2)

    return row_data


def save_results(rows, project_name):
    pd.DataFrame(rows).to_excel(get_project_results_file(project_name), index=False)


def load_saved_results(project_name):
    results_file = get_project_results_file(project_name)
    if results_file.exists():
        return pd.read_excel(results_file).fillna("").to_dict("records")
    return []


def save_full_project(project_name):
    final_name = safe_project_name(project_name)

    rows_to_save = []
    for row in st.session_state.rows:
        updated_row = dict(row)
        updated_row["Project Name"] = final_name
        updated_row["BOQ Article"] = st.session_state.boq_article
        rows_to_save.append(updated_row)

    save_results(rows_to_save, final_name)

    if "edited_supplier_df" in st.session_state and st.session_state.selected_supplier:
        save_df = st.session_state["edited_supplier_df"].copy()
        if not save_df.empty:
            save_df["Supplier"] = st.session_state.selected_supplier
        save_supplier_data_by_name(st.session_state.selected_supplier, save_df)

    st.session_state.project_name = final_name
    st.session_state.rows = rows_to_save


def open_full_project(project_name):
    final_name = safe_project_name(project_name)
    st.session_state.rows = load_saved_results(final_name)
    st.session_state.project_name = final_name

    if st.session_state.rows:
        st.session_state.boq_article = str(st.session_state.rows[0].get("BOQ Article", ""))
    else:
        st.session_state.boq_article = ""


df = load_profiles()
df["Profile Type"] = df["Profile"].astype(str).apply(get_profile_type)

profile_list = df["Profile"].dropna().astype(str).str.strip().tolist()
profile_type_options = sorted(df["Profile Type"].dropna().astype(str).unique().tolist())

floor_options = ["Ground Floor", "First Floor", "Second Floor", "Third Floor", "Fourth Floor"]
sub_article_options = ["Beam", "Column", "Brace", "Plate", "Connection"]

if "rows" not in st.session_state:
    st.session_state.rows = []
if "project_name" not in st.session_state:
    st.session_state.project_name = DEFAULT_PROJECT_NAME
if "boq_article" not in st.session_state:
    st.session_state.boq_article = ""
if "selected_supplier" not in st.session_state:
    st.session_state.selected_supplier = ""

main_tabs = st.tabs(["File", "Model", "Save", "Refresh", "Calculation"])

with main_tabs[0]:
    st.subheader("File")

    file_action = st.selectbox(
        "File Menu",
        ["Select", "New Project", "Open Project", "Import Project", "Export Project", "Rename Project"],
        key="file_menu_select_main"
    )

    if file_action == "New Project":
        c1, c2 = st.columns([2, 1])
        with c1:
            new_project_name = st.text_input("New Project Name", value=DEFAULT_PROJECT_NAME, key="file_new_project_name_input")
        with c2:
            st.write("")
            st.write("")
            if st.button("Create New Project", key="file_create_new_project_btn"):
                st.session_state.rows = []
                st.session_state.project_name = safe_project_name(new_project_name)
                st.session_state.boq_article = ""
                save_full_project(st.session_state.project_name)
                st.success(f"New project created and saved: {st.session_state.project_name}")
                st.rerun()

    elif file_action == "Open Project":
        existing_projects = get_saved_project_names()

        selected_project = st.selectbox(
            "Select Project",
            existing_projects if existing_projects else [DEFAULT_PROJECT_NAME],
            key="file_open_project_select"
        )

        if st.button("Open Selected Project", key="file_open_selected_project_btn"):
            open_full_project(selected_project)
            st.success(f"Opened: {st.session_state.project_name}")
            st.rerun()

    elif file_action == "Import Project":
        import_file = st.file_uploader("Import Excel Project", type=["xlsx"], key="file_import_project_uploader")
        if import_file is not None:
            imported_df = pd.read_excel(import_file).fillna("")
            st.session_state.rows = imported_df.to_dict("records")
            if st.session_state.rows:
                first_row = st.session_state.rows[0]
                st.session_state.project_name = str(first_row.get("Project Name", DEFAULT_PROJECT_NAME))
                st.session_state.boq_article = str(first_row.get("BOQ Article", ""))
            st.success("Project imported")

    elif file_action == "Export Project":
        if st.session_state.rows:
            export_df = pd.DataFrame(st.session_state.rows)
            st.download_button(
                label="Export Project",
                data=export_df.to_csv(index=False).encode("utf-8"),
                file_name=f"{safe_project_name(st.session_state.project_name)}.csv",
                mime="text/csv",
                key="file_export_project_btn"
            )
        else:
            st.info("No data to export")

    elif file_action == "Rename Project":
        c1, c2 = st.columns([2, 1])
        with c1:
            rename_to = st.text_input("Rename To", key="file_rename_to_input")
        with c2:
            st.write("")
            st.write("")
            if st.button("Rename Now", key="file_rename_now_btn"):
                if rename_to.strip():
                    old_name = st.session_state.project_name
                    new_name = safe_project_name(rename_to.strip())
                    rename_project_file(old_name, new_name)
                    st.session_state.project_name = new_name
                    for i in range(len(st.session_state.rows)):
                        st.session_state.rows[i]["Project Name"] = new_name
                    st.success(f"Project renamed to: {new_name}")
                    st.rerun()

with main_tabs[1]:
    st.subheader("Model")

    model_tabs = st.tabs([
        "Data Input",
        "Supplier Data",
        "Detail Results",
        "Summary by Floor",
        "Profile Sum",
        "Waste Calculation"
    ])

    recalculated_df = pd.DataFrame()
    summary_df = pd.DataFrame()
    profile_summary_df = pd.DataFrame()
    waste_df = pd.DataFrame()
    total_waste_weight = 0.0

    with model_tabs[0]:
        saved_projects = get_saved_project_names()
        project_options = saved_projects if saved_projects else [st.session_state.project_name]
        if st.session_state.project_name not in project_options:
            project_options = [st.session_state.project_name] + project_options

        current_index = project_options.index(st.session_state.project_name) if st.session_state.project_name in project_options else 0

        selected_model_project = st.selectbox(
            "Project Name",
            project_options,
            index=current_index,
            key="model_project_select_unique"
        )

        if selected_model_project != st.session_state.project_name:
            st.session_state.project_name = selected_model_project
            st.session_state.rows = load_saved_results(selected_model_project)

            if st.session_state.rows:
                st.session_state.boq_article = str(st.session_state.rows[0].get("BOQ Article", ""))
            else:
                st.session_state.boq_article = ""
            st.rerun()

        st.session_state.boq_article = st.text_input(
            "BOQ Article",
            value=st.session_state.boq_article,
            key="model_boq_article_input_unique"
        )

        c1, c2, c3 = st.columns(3)
        with c1:
            floor_level = st.selectbox("Floor Level", floor_options, key="model_floor_level_select")
        with c2:
            sub_article = st.selectbox("Sub Article", sub_article_options, key="model_sub_article_select")
        with c3:
            profile = st.selectbox("Profile", profile_list, key="model_profile_select")

        c4, c5, c6 = st.columns(3)
        with c4:
            input_length = st.number_input("Length (m)", min_value=0.0, step=0.1, format="%.2f", key="model_length_input")
        with c5:
            input_quantity = st.number_input("Quantity", min_value=1, step=1, key="model_quantity_input")
        with c6:
            input_price_per_ton = st.number_input("Price/t", min_value=0.0, step=10.0, format="%.2f", key="model_price_per_ton_input")

        current_data = {
            "Project Name": st.session_state.project_name,
            "BOQ Article": st.session_state.boq_article,
            "Floor Level": floor_level,
            "Sub Article": sub_article,
            "Profile": profile,
            "Length": input_length,
            "Quantity": input_quantity,
            "Price/t": input_price_per_ton,
            "Split Pieces": 1,
            "Split Length": 0.0,
            "Split Quantity": 0.0,
            "Total Treatment Area": 0.0,
            "Net Weight": 0.0,
            "Weight Incl. Waste": 0.0,
            "Total ZBSL": 0.0,
            "Total Levering Price": 0.0
        }

        current_data = calculate_row(current_data, df)

        r1, r2, r3 = st.columns(3)
        with r1:
            st.number_input("Split Pieces", value=int(to_float(current_data["Split Pieces"], 1)), disabled=True, key="model_calc_split_pieces_display")
        with r2:
            st.number_input("Net Weight", value=to_float(current_data["Net Weight"]), disabled=True, key="model_calc_net_weight_display")
        with r3:
            st.number_input("Weight Incl. Waste", value=to_float(current_data["Weight Incl. Waste"]), disabled=True, key="model_calc_weight_waste_display")

        r4, r5, r6 = st.columns(3)
        with r4:
            st.number_input("Treatment Area", value=to_float(current_data["Total Treatment Area"]), disabled=True, key="model_calc_treatment_area_display")
        with r5:
            st.number_input("ZBSL", value=to_float(current_data["Total ZBSL"]), disabled=True, key="model_calc_zbsl_display")
        with r6:
            st.number_input("Levering Price", value=to_float(current_data["Total Levering Price"]), disabled=True, key="model_calc_price_display")

        b1, b2 = st.columns(2)
        with b1:
            if st.button("Add Row", key="model_add_row_btn"):
                st.session_state.rows.append(current_data.copy())
                st.success("Row added")
        with b2:
            if st.button("Clear Rows", key="model_clear_rows_btn"):
                st.session_state.rows = []
                st.success("Rows cleared")

    with model_tabs[1]:
        st.subheader("Supplier Data")

        supplier_names = load_supplier_names()

        s0, s1 = st.columns([2, 1])
        with s0:
            selected_supplier_name = st.selectbox(
                "Select Supplier",
                supplier_names if supplier_names else [],
                key="supplier_select_supplier_box"
            ) if supplier_names else ""
        with s1:
            new_supplier_name = st.text_input("New Supplier Name", key="supplier_new_supplier_name_input")

        o1, o2 = st.columns(2)
        with o1:
            if st.button("Open Supplier", key="supplier_open_supplier_btn"):
                if selected_supplier_name:
                    st.session_state.selected_supplier = selected_supplier_name
                    st.success(f"Opened supplier: {selected_supplier_name}")
                    st.rerun()

        with o2:
            if st.button("Create Supplier", key="supplier_create_supplier_btn"):
                if new_supplier_name.strip():
                    supplier_name = safe_name(new_supplier_name.strip())
                    empty_df = pd.DataFrame(columns=["Supplier", "Profile Type", "Fabric Standard Length"])
                    save_supplier_data_by_name(supplier_name, empty_df)
                    st.session_state.selected_supplier = supplier_name
                    st.success(f"Created supplier: {supplier_name}")
                    st.rerun()

        active_supplier = st.session_state.get("selected_supplier", "")

        if active_supplier:
            st.text_input("Active Supplier", value=active_supplier, disabled=True, key="supplier_active_supplier_display_unique")

            supplier_df = load_supplier_data_by_name(active_supplier)

            s2, s3, s4 = st.columns(3)
            with s2:
                selected_profile_type = st.selectbox("Supplier Profile Type", profile_type_options, key="supplier_profile_type_select")
            with s3:
                fabric_standard_length_input = st.number_input("Supplier Fabric Standard Length", min_value=0.0, step=0.5, key="supplier_fabric_length_input")
            with s4:
                st.write("")
                st.write("")
                if st.button("Add Supplier Data", key="supplier_add_supplier_data_btn"):
                    new_row = pd.DataFrame([{
                        "Supplier": active_supplier,
                        "Profile Type": selected_profile_type,
                        "Fabric Standard Length": fabric_standard_length_input
                    }])

                    supplier_df = supplier_df[
                        ~(supplier_df["Profile Type"].astype(str).str.strip() == selected_profile_type)
                    ]

                    supplier_df = pd.concat([supplier_df, new_row], ignore_index=True)
                    save_supplier_data_by_name(active_supplier, supplier_df)
                    st.success("Supplier data saved")
                    st.rerun()

            edited_supplier_df = st.data_editor(
                supplier_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key="supplier_editor_data_editor"
            )
            st.session_state["edited_supplier_df"] = edited_supplier_df.copy()
        else:
            st.info("Create or open a supplier first.")

    if st.session_state.rows:
        detail_df = pd.DataFrame(st.session_state.rows)

        expected_columns = [
            "Project Name", "BOQ Article", "Floor Level", "Sub Article", "Profile",
            "Length", "Quantity", "Price/t", "Split Pieces",
            "Split Length", "Split Quantity",
            "Total Treatment Area", "Net Weight", "Weight Incl. Waste",
            "Total ZBSL", "Total Levering Price"
        ]

        for col in expected_columns:
            if col not in detail_df.columns:
                detail_df[col] = 0

        detail_df = detail_df[expected_columns]

        with model_tabs[2]:
            st.subheader("Detail Results")

            detail_display_df = detail_df[[
                "Project Name", "BOQ Article", "Floor Level", "Sub Article", "Profile",
                "Length", "Quantity", "Price/t", "Split Pieces",
                "Total Treatment Area", "Net Weight", "Weight Incl. Waste",
                "Total ZBSL", "Total Levering Price"
            ]]

            edited_df = st.data_editor(
                detail_display_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key="detail_results_editor",
                column_config={
                    "Length": st.column_config.NumberColumn("Length", step=0.1, format="%.2f"),
                    "Quantity": st.column_config.NumberColumn("Quantity", step=1),
                    "Price/t": st.column_config.NumberColumn("Price/t", step=1.0, format="%.2f"),
                    "Split Pieces": st.column_config.NumberColumn("Split Pieces", disabled=True),
                    "Total Treatment Area": st.column_config.NumberColumn("Total Treatment Area", disabled=True, format="%.2f"),
                    "Net Weight": st.column_config.NumberColumn("Net Weight", disabled=True, format="%.2f"),
                    "Weight Incl. Waste": st.column_config.NumberColumn("Weight Incl. Waste", disabled=True, format="%.2f"),
                    "Total ZBSL": st.column_config.NumberColumn("Total ZBSL", disabled=True, format="%.2f"),
                    "Total Levering Price": st.column_config.NumberColumn("Total Levering Price", disabled=True, format="%.2f"),
                }
            )

            auto_recalc_rows = []
            for _, row_item in edited_df.iterrows():
                row_dict = row_item.to_dict()
                row_dict["Project Name"] = st.session_state.project_name
                row_dict["BOQ Article"] = st.session_state.boq_article
                row_dict = calculate_row(row_dict, df)
                auto_recalc_rows.append(row_dict)

            st.session_state.rows = auto_recalc_rows
            recalculated_df = pd.DataFrame(auto_recalc_rows).fillna(0)

        if not recalculated_df.empty:
            summary_df = recalculated_df.groupby(
                ["Floor Level", "Sub Article"], as_index=False
            )[["Quantity", "Split Quantity", "Total Treatment Area", "Net Weight", "Weight Incl. Waste", "Total ZBSL", "Total Levering Price"]].sum()

            summary_df = summary_df.rename(columns={
                "Quantity": "Input Quantity",
                "Split Quantity": "Calculated Quantity"
            })

            profile_summary_df = recalculated_df.groupby("Profile", as_index=False).agg({
                "Length": "sum",
                "Split Length": "sum",
                "Quantity": "sum",
                "Split Quantity": "sum",
                "Weight Incl. Waste": "sum"
            })

            profile_summary_df = profile_summary_df.rename(columns={
                "Length": "Total Input Length",
                "Split Length": "Total Calculated Length",
                "Quantity": "Total Input Quantity",
                "Split Quantity": "Total Calculated Quantity",
                "Weight Incl. Waste": "Total Weight"
            })

            active_supplier = st.session_state.get("selected_supplier", "")
            supplier_df = load_supplier_data_by_name(active_supplier) if active_supplier else pd.DataFrame()

            if active_supplier and not supplier_df.empty:
                waste_df = profile_summary_df.copy()
                waste_df["Profile Type"] = waste_df["Profile"].apply(get_profile_type)

                waste_df["Fabric Standard Length"] = waste_df["Profile Type"].apply(
                    lambda pt: to_float(
                        get_supplier_row(pt, supplier_df)["Fabric Standard Length"],
                        0.0
                    ) if get_supplier_row(pt, supplier_df) is not None else 0.0
                )

                waste_df["Supplier Qty"] = waste_df.apply(
                    lambda row: math.ceil(row["Total Calculated Length"] / row["Fabric Standard Length"])
                    if to_float(row["Fabric Standard Length"]) > 0 else 0,
                    axis=1
                )

                waste_df["Waste Length"] = waste_df.apply(
                    lambda row: round(
                        row["Supplier Qty"] * row["Fabric Standard Length"] - row["Total Calculated Length"], 2
                    ) if to_float(row["Fabric Standard Length"]) > 0 else 0.0,
                    axis=1
                )

                kgm_map = df.set_index("Profile")["kgm"].to_dict()
                waste_df["Waste Weight"] = waste_df.apply(
                    lambda row: round(row["Waste Length"] * to_float(kgm_map.get(row["Profile"], 0)), 2),
                    axis=1
                )

                waste_df = waste_df[
                    ["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"]
                ].fillna(0)
                total_waste_weight = round(waste_df["Waste Weight"].sum(), 2)

        with model_tabs[3]:
            st.subheader("Summary by Floor")
            if not summary_df.empty:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
            else:
                st.info("No summary yet")

        with model_tabs[4]:
            st.subheader("Profile Sum")
            if not profile_summary_df.empty:
                st.dataframe(profile_summary_df, use_container_width=True, hide_index=True)
            else:
                st.info("No profile sum yet")

        with model_tabs[5]:
            st.subheader("Waste Calculation")
            active_supplier = st.session_state.get("selected_supplier", "")
            if active_supplier:
                st.text_input("Waste Supplier", value=active_supplier, disabled=True, key="waste_supplier_display_unique")

            if not waste_df.empty:
                total_row = pd.DataFrame([{
                    "Profile": "",
                    "Fabric Standard Length": "",
                    "Supplier Qty": "",
                    "Waste Length": "Total",
                    "Waste Weight": total_waste_weight
                }])

                waste_display = pd.concat([waste_df, total_row], ignore_index=True)
                st.dataframe(waste_display, use_container_width=True, hide_index=True)
            else:
                st.info("Open a supplier and add supplier data first.")
    else:
        with model_tabs[2]:
            st.subheader("Detail Results")
            st.info("No detail data")
        with model_tabs[3]:
            st.subheader("Summary by Floor")
            st.info("No summary yet")
        with model_tabs[4]:
            st.subheader("Profile Sum")
            st.info("No profile sum yet")
        with model_tabs[5]:
            st.subheader("Waste Calculation")
            st.info("Open a supplier and add supplier data first.")

with main_tabs[2]:
    st.subheader("Save")

    save_project_name = st.text_input(
        "Project Name",
        value=st.session_state.project_name,
        key="save_project_name_input_unique"
    )

    if st.button("Save", use_container_width=True, key="save_project_btn_unique"):
        final_name = save_project_name.strip() if save_project_name.strip() else DEFAULT_PROJECT_NAME
        save_full_project(final_name)
        st.success(f"Project saved: {st.session_state.project_name}")
        st.rerun()

with main_tabs[3]:
    st.subheader("Refresh")

    if st.button("Refresh", use_container_width=True, key="refresh_app_btn_unique"):
        st.rerun()

with main_tabs[4]:
    st.subheader("Calculation")

    calc_action = st.selectbox(
        "Calculation Menu",
        ["Select", "Connection", "Bolt", "Weld", "Plate", "Custom Code"],
        key="calc_menu_select_unique"
    )

    if calc_action != "Select":
        st.info(f"{calc_action} section is ready for future code")

if st.session_state.rows:
    export_detail_df = pd.DataFrame(st.session_state.rows).fillna(0)

    export_summary_df = export_detail_df.groupby(
        ["Floor Level", "Sub Article"], as_index=False
    )[["Quantity", "Split Quantity", "Total Treatment Area", "Net Weight", "Weight Incl. Waste", "Total ZBSL", "Total Levering Price"]].sum()

    export_summary_df = export_summary_df.rename(columns={
        "Quantity": "Input Quantity",
        "Split Quantity": "Calculated Quantity"
    })

    export_profile_sum_df = export_detail_df.groupby("Profile", as_index=False).agg({
        "Length": "sum",
        "Split Length": "sum",
        "Quantity": "sum",
        "Split Quantity": "sum",
        "Weight Incl. Waste": "sum"
    }).rename(columns={
        "Length": "Total Input Length",
        "Split Length": "Total Calculated Length",
        "Quantity": "Total Input Quantity",
        "Split Quantity": "Total Calculated Quantity",
        "Weight Incl. Waste": "Total Weight"
    })

    export_waste_df = pd.DataFrame(columns=["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"])
    export_total_waste_weight = 0.0

    active_supplier = st.session_state.get("selected_supplier", "")
    export_supplier_df = load_supplier_data_by_name(active_supplier) if active_supplier else pd.DataFrame()

    if active_supplier and not export_supplier_df.empty:
        export_waste_df = export_profile_sum_df.copy()
        export_waste_df["Profile Type"] = export_waste_df["Profile"].apply(get_profile_type)

        export_waste_df["Fabric Standard Length"] = export_waste_df["Profile Type"].apply(
            lambda pt: to_float(
                get_supplier_row(pt, export_supplier_df)["Fabric Standard Length"],
                0.0
            ) if get_supplier_row(pt, export_supplier_df) is not None else 0.0
        )

        export_waste_df["Supplier Qty"] = export_waste_df.apply(
            lambda row: math.ceil(row["Total Calculated Length"] / row["Fabric Standard Length"])
            if to_float(row["Fabric Standard Length"]) > 0 else 0,
            axis=1
        )

        export_waste_df["Waste Length"] = export_waste_df.apply(
            lambda row: round(
                row["Supplier Qty"] * row["Fabric Standard Length"] - row["Total Calculated Length"], 2
            ) if to_float(row["Fabric Standard Length"]) > 0 else 0.0,
            axis=1
        )

        kgm_map = df.set_index("Profile")["kgm"].to_dict()
        export_waste_df["Waste Weight"] = export_waste_df.apply(
            lambda row: round(row["Waste Length"] * to_float(kgm_map.get(row["Profile"], 0)), 2),
            axis=1
        )

        export_waste_df = export_waste_df[
            ["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"]
        ].fillna(0)

        export_total_waste_weight = round(export_waste_df["Waste Weight"].sum(), 2)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_detail_df.to_excel(writer, index=False, sheet_name="Detail Results")
        export_summary_df.to_excel(writer, index=False, sheet_name="Summary by Floor")
        export_profile_sum_df.to_excel(writer, index=False, sheet_name="Profile Sum")
        export_waste_df.to_excel(writer, index=False, sheet_name="Waste Calculation")

        total_waste_df = pd.DataFrame({
            "Item": ["Total Waste Weight"],
            "Value": [export_total_waste_weight]
        })
        total_waste_df.to_excel(writer, index=False, sheet_name="Total Waste")

    output.seek(0)

    st.download_button(
        label="Export to Excel",
        data=output,
        file_name=f"{safe_project_name(st.session_state.project_name)}_steel_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="export_excel_btn_unique"
    )
