import streamlit as st
import pandas as pd
import io
import os
import math
import re
from pathlib import Path

st.set_page_config(page_title="SEST", layout="wide")
st.title("SEST")

PROJECTS_DIR = Path("projects")
PROJECTS_DIR.mkdir(exist_ok=True)

PROFILES_FILE = "Profiles.xlsx"
MAX_PIECE_LENGTH = 23.0
DEFAULT_PROJECT_NAME = "default_project"


def safe_project_name(name):
    name = str(name).strip()
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", name)
    return name if name else DEFAULT_PROJECT_NAME


def get_project_results_file(project_name):
    return PROJECTS_DIR / f"{safe_project_name(project_name)}_results.xlsx"


def get_project_supplier_file(project_name):
    return PROJECTS_DIR / f"{safe_project_name(project_name)}_supplier.xlsx"


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

    if p.startswith(("HEA", "HEB", "HEM", "IPE", "IPN", "UPN")):
        return 1.15
    elif p.startswith(("K", "L", "R")):
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


def load_supplier_data(project_name):
    supplier_file = get_project_supplier_file(project_name)

    if supplier_file.exists():
        df = pd.read_excel(supplier_file).fillna("")
        expected_cols = ["Supplier", "Profile Type", "Fabric Standard Length"]
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ""
        return df[expected_cols]

    return pd.DataFrame(columns=["Supplier", "Profile Type", "Fabric Standard Length"])


def save_supplier_data(df, project_name):
    supplier_file = get_project_supplier_file(project_name)
    df.to_excel(supplier_file, index=False)


def get_supplier_row(profile_type, supplier_name, supplier_df):
    match = supplier_df[
        (supplier_df["Profile Type"].astype(str).str.strip() == str(profile_type).strip()) &
        (supplier_df["Supplier"].astype(str).str.strip() == str(supplier_name).strip())
    ]
    if not match.empty:
        return match.iloc[0]
    return None


def rename_project_files(old_name, new_name):
    old_results = get_project_results_file(old_name)
    old_supplier = get_project_supplier_file(old_name)

    new_results = get_project_results_file(new_name)
    new_supplier = get_project_supplier_file(new_name)

    if old_results.exists():
        old_results.rename(new_results)

    if old_supplier.exists():
        old_supplier.rename(new_supplier)


def calculate_row(row_data, profile_df):
    profile_name = str(row_data.get("Profile", "")).strip()

    default_result = {
        "Split Pieces": 1,
        "kg/m": 0.0,
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

    original_length = to_float(row_data.get("Length", 0))
    original_number = to_float(row_data.get("Number", 0))
    price_per_ton = to_float(row_data.get("Price/t", 0))

    calc_length, calc_number, split_pieces = split_length_and_quantity(
        original_length, original_number, MAX_PIECE_LENGTH
    )

    kgm = to_float(profile_row.get("kgm", 0))
    m2_per_m = to_float(profile_row.get("m2_per_m", 0))
    zbsl = get_zbsl(profile_row, calc_length)

    net_weight = kgm * calc_length * calc_number
    factor = get_weight_factor(profile_name)
    weight_incl_waste = net_weight * factor

    total_treatment_area = m2_per_m * calc_length * calc_number
    total_zbsl = zbsl * calc_number
    total_price = (weight_incl_waste / 1000) * price_per_ton

    row_data["Length"] = round(calc_length, 2)
    row_data["Number"] = int(calc_number) if float(calc_number).is_integer() else round(calc_number, 2)
    row_data["Split Pieces"] = int(split_pieces)
    row_data["Price/t"] = round(price_per_ton, 2)
    row_data["kg/m"] = round(kgm, 2)
    row_data["Total Treatment Area"] = round(total_treatment_area, 2)
    row_data["Net Weight"] = round(net_weight, 2)
    row_data["Weight Incl. Waste"] = round(weight_incl_waste, 2)
    row_data["Total ZBSL"] = round(total_zbsl, 2)
    row_data["Total Levering Price"] = round(total_price, 2)

    return row_data


def save_results(rows, project_name):
    results_file = get_project_results_file(project_name)
    pd.DataFrame(rows).to_excel(results_file, index=False)


def load_saved_results(project_name):
    results_file = get_project_results_file(project_name)
    if results_file.exists():
        saved_df = pd.read_excel(results_file).fillna("")
        return saved_df.to_dict("records")
    return []


df = load_profiles()
df["Profile Type"] = df["Profile"].astype(str).apply(get_profile_type)

profile_list = df["Profile"].dropna().astype(str).str.strip().tolist()
profile_type_options = sorted(df["Profile Type"].dropna().astype(str).unique().tolist())

floor_options = ["Ground Floor", "First Floor", "Second Floor", "Third Floor", "Fourth Floor"]
sub_article_options = ["Beam", "Column", "Brace", "Plate", "Connection"]

if "active_project" not in st.session_state:
    st.session_state.active_project = ""

if "rows" not in st.session_state:
    st.session_state.rows = []

if "loaded_boq" not in st.session_state:
    st.session_state.loaded_boq = ""

st.subheader("Project")

existing_projects = sorted([
    p.name.replace("_results.xlsx", "")
    for p in PROJECTS_DIR.glob("*_results.xlsx")
])

p1, p2 = st.columns([1, 2])

with p1:
    project_mode = st.radio("Mode", ["New Project", "Open Project", "Save"])

with p2:
    project = ""
    rename_project = ""

    if project_mode == "New Project":
        project_input = st.text_input("Project Name")
        project = project_input.strip() if project_input.strip() else DEFAULT_PROJECT_NAME

        if st.session_state.active_project != DEFAULT_PROJECT_NAME:
            st.session_state.active_project = DEFAULT_PROJECT_NAME
            st.session_state.rows = []
            st.session_state.loaded_boq = ""

    elif project_mode == "Open Project":
        project = st.selectbox("Select Project", existing_projects) if existing_projects else ""
        if st.button("Load Project"):
            if project:
                st.session_state.active_project = project
                st.session_state.rows = load_saved_results(project)
                if st.session_state.rows:
                    first_row = st.session_state.rows[0]
                    st.session_state.loaded_boq = str(first_row.get("BOQ Article", ""))
                else:
                    st.session_state.loaded_boq = ""
                st.success(f"Loaded: {project}")

    elif project_mode == "Save":
        save_project_name = st.text_input(
            "Save Project Name",
            value=st.session_state.active_project if st.session_state.active_project else ""
        ).strip()

        rename_project = st.text_input("Rename To", key="rename_project")

        sp1, sp2 = st.columns(2)

        with sp1:
            if st.button("Save Project"):
                save_name = save_project_name if save_project_name else DEFAULT_PROJECT_NAME
                rows_to_save = []

                for row in st.session_state.rows:
                    updated_row = dict(row)
                    updated_row["Project Name"] = save_name
                    updated_row["BOQ Article"] = st.session_state.loaded_boq
                    rows_to_save.append(updated_row)

                save_results(rows_to_save, save_name)

                current_supplier_df = load_supplier_data(st.session_state.active_project or DEFAULT_PROJECT_NAME)
                save_supplier_data(current_supplier_df, save_name)

                st.session_state.active_project = save_name
                st.session_state.rows = rows_to_save
                st.success(f"Saved: {save_name}")

        with sp2:
            if st.button("Rename Project"):
                if (
                    st.session_state.active_project
                    and rename_project.strip()
                    and st.session_state.active_project != DEFAULT_PROJECT_NAME
                ):
                    rename_project_files(st.session_state.active_project, rename_project.strip())
                    st.session_state.active_project = rename_project.strip()
                    for i in range(len(st.session_state.rows)):
                        st.session_state.rows[i]["Project Name"] = rename_project.strip()
                    st.success(f"Renamed to: {rename_project.strip()}")
                    st.rerun()

boq_default = st.session_state.loaded_boq if project_mode != "New Project" else ""
boq = st.text_input("BOQ Article", value=boq_default)

if boq != st.session_state.loaded_boq:
    st.session_state.loaded_boq = boq

left_col, right_col = st.columns([2, 1])

with left_col:
    st.subheader("Input Data")

    i1, i2, i3 = st.columns(3)
    with i1:
        floor_level = st.selectbox("Floor Level", floor_options)
    with i2:
        sub_article = st.selectbox("Sub Article", sub_article_options)
    with i3:
        profile = st.selectbox("Profile", profile_list)

    i4, i5, i6 = st.columns(3)
    with i4:
        length = st.number_input("Length (m)", min_value=0.0, step=0.1, format="%.2f")
    with i5:
        quantity = st.number_input("Quantity", min_value=1, step=1)
    with i6:
        price_per_ton = st.number_input("Price per ton", min_value=0.0, step=10.0, format="%.2f")

    current_project_name = st.session_state.active_project or DEFAULT_PROJECT_NAME

    current_data = {
        "Project Name": current_project_name,
        "BOQ Article": boq,
        "Floor Level": floor_level,
        "Sub Article": sub_article,
        "Profile": profile,
        "Length": length,
        "Number": quantity,
        "Price/t": price_per_ton,
        "Split Pieces": 1,
        "kg/m": 0.0,
        "Total Treatment Area": 0.0,
        "Net Weight": 0.0,
        "Weight Incl. Waste": 0.0,
        "Total ZBSL": 0.0,
        "Total Levering Price": 0.0
    }

    current_data = calculate_row(current_data, df)

    b1, b2 = st.columns(2)
    with b1:
        if st.button("Add"):
            st.session_state.rows.append(current_data.copy())
            st.success("Row added.")
    with b2:
        if st.button("Clear Screen"):
            st.session_state.rows = []
            st.success("Current screen cleared.")

with right_col:
    st.subheader("Supplier Data")

    supplier_df = load_supplier_data(st.session_state.active_project or DEFAULT_PROJECT_NAME)

    supplier_name_input = st.text_input("Supplier").strip()
    selected_profile_type = st.selectbox("Profile Type", profile_type_options)
    fabric_standard_length_input = st.number_input("Fabric Standard Length", min_value=0.0, step=0.5)

    if st.button("Add Supplier Data"):
        if supplier_name_input == "":
            st.warning("Please enter a Supplier name.")
        else:
            new_row = pd.DataFrame([{
                "Supplier": supplier_name_input,
                "Profile Type": selected_profile_type,
                "Fabric Standard Length": fabric_standard_length_input
            }])

            supplier_df = supplier_df[
                ~(
                    (supplier_df["Supplier"].astype(str).str.strip() == supplier_name_input) &
                    (supplier_df["Profile Type"].astype(str).str.strip() == selected_profile_type)
                )
            ]

            supplier_df = pd.concat([supplier_df, new_row], ignore_index=True)
            save_supplier_data(supplier_df, st.session_state.active_project or DEFAULT_PROJECT_NAME)
            st.success("Supplier data saved.")
            supplier_df = load_supplier_data(st.session_state.active_project or DEFAULT_PROJECT_NAME)

    if not supplier_df.empty:
        st.dataframe(supplier_df, use_container_width=True, hide_index=True)

st.subheader("Detail Results")

if st.session_state.rows:
    detail_df = pd.DataFrame(st.session_state.rows)

    expected_columns = [
        "Project Name", "BOQ Article", "Floor Level", "Sub Article", "Profile",
        "Length", "Number", "Price/t", "Split Pieces", "kg/m",
        "Total Treatment Area", "Net Weight", "Weight Incl. Waste",
        "Total ZBSL", "Total Levering Price"
    ]

    for col in expected_columns:
        if col not in detail_df.columns:
            detail_df[col] = 0

    detail_df = detail_df[expected_columns]

    edited_df = st.data_editor(
        detail_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="detail_editor",
        column_config={
            "Project Name": st.column_config.TextColumn("Project Name"),
            "BOQ Article": st.column_config.TextColumn("BOQ Article"),
            "Floor Level": st.column_config.SelectboxColumn("Floor Level", options=floor_options),
            "Sub Article": st.column_config.SelectboxColumn("Sub Article", options=sub_article_options),
            "Profile": st.column_config.SelectboxColumn("Profile", options=profile_list),
            "Length": st.column_config.NumberColumn("Length", step=0.1, format="%.2f"),
            "Number": st.column_config.NumberColumn("Number", step=1),
            "Price/t": st.column_config.NumberColumn("Price/t", step=10.0, format="%.2f"),
            "Split Pieces": st.column_config.NumberColumn("Split Pieces", disabled=True),
            "kg/m": st.column_config.NumberColumn("kg/m", disabled=True, format="%.2f"),
            "Total Treatment Area": st.column_config.NumberColumn("Total Treatment Area", disabled=True, format="%.2f"),
            "Net Weight": st.column_config.NumberColumn("Net Weight", disabled=True, format="%.2f"),
            "Weight Incl. Waste": st.column_config.NumberColumn("Weight Incl. Waste", disabled=True, format="%.2f"),
            "Total ZBSL": st.column_config.NumberColumn("Total ZBSL", disabled=True, format="%.2f"),
            "Total Levering Price": st.column_config.NumberColumn("Total Levering Price", disabled=True, format="%.2f"),
        }
    )

    recalculated_rows = []
    for _, row_item in edited_df.iterrows():
        row_dict = row_item.to_dict()
        row_dict["Project Name"] = row_dict.get("Project Name", st.session_state.active_project or DEFAULT_PROJECT_NAME)
        row_dict["BOQ Article"] = row_dict.get("BOQ Article", st.session_state.loaded_boq)
        row_dict = calculate_row(row_dict, df)
        recalculated_rows.append(row_dict)

    recalculated_df = pd.DataFrame(recalculated_rows).fillna(0)
    st.session_state.rows = recalculated_df.to_dict("records")

    st.subheader("Summary by Floor")
    summary_df = recalculated_df.groupby(
        ["Floor Level", "Sub Article"], as_index=False
    )[["Number", "Total Treatment Area", "Net Weight", "Weight Incl. Waste", "Total ZBSL", "Total Levering Price"]].sum()
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.subheader("Profile Sum")
    profile_summary_df = recalculated_df.groupby("Profile", as_index=False).agg({
        "Length": "sum",
        "Number": "sum",
        "Weight Incl. Waste": "sum"
    })

    profile_summary_df = profile_summary_df.rename(columns={
        "Length": "Total Length",
        "Number": "Total Number",
        "Weight Incl. Waste": "Total Weight"
    })

    st.dataframe(profile_summary_df, use_container_width=True, hide_index=True)

    st.subheader("Waste Calculation")

    supplier_df = load_supplier_data(st.session_state.active_project or DEFAULT_PROJECT_NAME)
    supplier_options = sorted(supplier_df["Supplier"].dropna().astype(str).str.strip().unique().tolist())

    if supplier_options:
        selected_supplier = st.selectbox("Select Supplier", supplier_options)

        waste_df = profile_summary_df.copy()
        waste_df["Profile Type"] = waste_df["Profile"].apply(get_profile_type)

        waste_df["Fabric Standard Length"] = waste_df["Profile Type"].apply(
            lambda pt: to_float(
                get_supplier_row(pt, selected_supplier, supplier_df)["Fabric Standard Length"],
                0.0
            ) if get_supplier_row(pt, selected_supplier, supplier_df) is not None else 0.0
        )

        waste_df["Supplier Qty"] = waste_df.apply(
            lambda row: math.ceil(row["Total Length"] / row["Fabric Standard Length"])
            if to_float(row["Fabric Standard Length"]) > 0 else 0,
            axis=1
        )

        waste_df["kg/m"] = waste_df["Profile"].apply(
            lambda p: to_float(
                df[df["Profile"].astype(str).str.strip() == str(p).strip()].iloc[0]["kgm"],
                0.0
            ) if not df[df["Profile"].astype(str).str.strip() == str(p).strip()].empty else 0.0
        )

        waste_df["Waste Length"] = waste_df.apply(
            lambda row: round(
                row["Supplier Qty"] * row["Fabric Standard Length"] - row["Total Length"], 2
            ) if to_float(row["Fabric Standard Length"]) > 0 else 0.0,
            axis=1
        )

        waste_df["Waste Weight"] = waste_df.apply(
            lambda row: round(row["Waste Length"] * row["kg/m"], 2),
            axis=1
        )

        waste_df = waste_df[
            ["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"]
        ].fillna(0)

        total_waste_weight = round(waste_df["Waste Weight"].sum(), 2)

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
        waste_df = pd.DataFrame(columns=["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"])
        total_waste_weight = 0.0
        st.info("Add Supplier Data first.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        recalculated_df.to_excel(writer, index=False, sheet_name="Detail Results")
        summary_df.to_excel(writer, index=False, sheet_name="Summary by Floor")
        profile_summary_df.to_excel(writer, index=False, sheet_name="Profile Sum")
        waste_df.to_excel(writer, index=False, sheet_name="Waste Calculation")

        total_waste_df = pd.DataFrame({
            "Item": ["Total Waste Weight"],
            "Value": [total_waste_weight]
        })
        total_waste_df.to_excel(writer, index=False, sheet_name="Total Waste")

    output.seek(0)

    export_name = st.session_state.active_project or DEFAULT_PROJECT_NAME

    st.download_button(
        label="Export to Excel",
        data=output,
        file_name=f"{safe_project_name(export_name)}_steel_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("No rows added yet.")
