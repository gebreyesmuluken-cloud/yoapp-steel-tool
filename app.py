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


def export_project_excel(project_name, detail_df, summary_df, profile_summary_df, waste_df, total_waste_weight):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Detail Results")
        summary_df.to_excel(writer, index=False, sheet_name="Summary by Floor")
        profile_summary_df.to_excel(writer, index=False, sheet_name="Profile Sum")
        waste_df.to_excel(writer, index=False, sheet_name="Waste Calculation")

        total_waste_df = pd.DataFrame({
            "Item": ["Total Waste Weight"],
            "Value": [total_waste_weight]
        })
        total_waste_df.to_excel(writer, index=False, sheet_name="Total Waste")

    output.seek(0)
    return output


df = load_profiles()
df["Profile Type"] = df["Profile"].astype(str).apply(get_profile_type)

profile_list = df["Profile"].dropna().astype(str).str.strip().tolist()
profile_type_options = sorted(df["Profile Type"].dropna().astype(str).unique().tolist())

floor_options = ["Ground Floor", "First Floor", "Second Floor", "Third Floor", "Fourth Floor"]
sub_article_options = ["Beam", "Column", "Brace", "Plate", "Connection"]

if "active_project" not in st.session_state:
    st.session_state.active_project = DEFAULT_PROJECT_NAME

if "rows" not in st.session_state:
    st.session_state.rows = []

if "loaded_boq" not in st.session_state:
    st.session_state.loaded_boq = ""

if "edit_mode" not in st.session_state:
    st.session_state.edit_mode = False

if "menu_main" not in st.session_state:
    st.session_state.menu_main = "Model"

if "menu_sub" not in st.session_state:
    st.session_state.menu_sub = "Data Input"

toolbar_cols = st.columns([1, 1, 1, 1, 4])

with toolbar_cols[0]:
    if st.button("File", use_container_width=True):
        st.session_state.menu_main = "File"

with toolbar_cols[1]:
    if st.button("Model", use_container_width=True):
        st.session_state.menu_main = "Model"

with toolbar_cols[2]:
    if st.button("Edit", use_container_width=True):
        st.session_state.menu_main = "Edit"

with toolbar_cols[3]:
    if st.button("Calculation", use_container_width=True):
        st.session_state.menu_main = "Calculation"

st.markdown("---")

if st.session_state.menu_main == "File":
    st.subheader("File")

    file_cols = st.columns(5)

    with file_cols[0]:
        if st.button("New Project", use_container_width=True):
            st.session_state.active_project = DEFAULT_PROJECT_NAME
            st.session_state.rows = []
            st.session_state.loaded_boq = ""
            st.success("New project ready.")

    with file_cols[1]:
        existing_projects = sorted([
            p.name.replace("_results.xlsx", "")
            for p in PROJECTS_DIR.glob("*_results.xlsx")
        ])
        selected_open_project = st.selectbox("Open Project", existing_projects if existing_projects else [DEFAULT_PROJECT_NAME])

    with file_cols[2]:
        if st.button("Load", use_container_width=True):
            st.session_state.active_project = selected_open_project
            st.session_state.rows = load_saved_results(selected_open_project)
            if st.session_state.rows:
                first_row = st.session_state.rows[0]
                st.session_state.loaded_boq = str(first_row.get("BOQ Article", ""))
            else:
                st.session_state.loaded_boq = ""
            st.success(f"Loaded: {selected_open_project}")

    with file_cols[3]:
        import_file = st.file_uploader("Import Project", type=["xlsx"])

    with file_cols[4]:
        project_save_name = st.text_input("Project Name", value=st.session_state.active_project)

    if import_file is not None:
        imported_df = pd.read_excel(import_file).fillna("")
        st.session_state.rows = imported_df.to_dict("records")
        st.success("Project imported.")

    save_cols = st.columns(3)

    with save_cols[0]:
        if st.button("Save Project", use_container_width=True):
            save_name = project_save_name.strip() if project_save_name.strip() else DEFAULT_PROJECT_NAME
            rows_to_save = []
            for row in st.session_state.rows:
                new_row = dict(row)
                new_row["Project Name"] = save_name
                new_row["BOQ Article"] = st.session_state.loaded_boq
                rows_to_save.append(new_row)

            save_results(rows_to_save, save_name)

            supplier_df_to_save = load_supplier_data(st.session_state.active_project or DEFAULT_PROJECT_NAME)
            save_supplier_data(supplier_df_to_save, save_name)

            st.session_state.active_project = save_name
            st.session_state.rows = rows_to_save
            st.success(f"Saved: {save_name}")

    with save_cols[1]:
        rename_to = st.text_input("Rename To")
        if st.button("Rename Project", use_container_width=True):
            if rename_to.strip():
                rename_project_files(st.session_state.active_project, rename_to.strip())
                st.session_state.active_project = rename_to.strip()
                st.success(f"Renamed to: {rename_to.strip()}")

    with save_cols[2]:
        if st.session_state.rows:
            current_df = pd.DataFrame(st.session_state.rows).fillna("")
            if not current_df.empty:
                summary_df = current_df.groupby(
                    ["Floor Level", "Sub Article"], as_index=False
                )[["Number", "Total Treatment Area", "Net Weight", "Weight Incl. Waste", "Total ZBSL", "Total Levering Price"]].sum()

                profile_summary_df = current_df.groupby("Profile", as_index=False).agg({
                    "Length": "sum",
                    "Number": "sum",
                    "Weight Incl. Waste": "sum"
                }).rename(columns={
                    "Length": "Total Length",
                    "Number": "Total Number",
                    "Weight Incl. Waste": "Total Weight"
                })

                waste_df = pd.DataFrame(columns=["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"])
                total_waste_weight = 0.0

                export_output = export_project_excel(
                    st.session_state.active_project,
                    current_df,
                    summary_df,
                    profile_summary_df,
                    waste_df,
                    total_waste_weight
                )

                st.download_button(
                    label="Export Project",
                    data=export_output,
                    file_name=f"{safe_project_name(st.session_state.active_project)}_project_export.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

if st.session_state.menu_main == "Model":
    st.subheader("Model")

    model_tabs = st.tabs([
        "Data Input",
        "Supplier Input",
        "Detail Results",
        "Summary by Floor",
        "Profile Sum",
        "Waste Calculation"
    ])

    with model_tabs[0]:
        top_cols = st.columns(2)
        with top_cols[0]:
            st.text_input("Project Name", value=st.session_state.active_project, disabled=True)
        with top_cols[1]:
            boq = st.text_input("BOQ Article", value=st.session_state.loaded_boq)
            st.session_state.loaded_boq = boq

        input_cols_1 = st.columns(3)
        with input_cols_1[0]:
            floor_level = st.selectbox("Floor Level", floor_options)
        with input_cols_1[1]:
            sub_article = st.selectbox("Sub Article", sub_article_options)
        with input_cols_1[2]:
            profile = st.selectbox("Profile", profile_list)

        input_cols_2 = st.columns(3)
        with input_cols_2[0]:
            length = st.number_input("Length (m)", min_value=0.0, step=0.1, format="%.2f")
        with input_cols_2[1]:
            quantity = st.number_input("Quantity", min_value=1, step=1)
        with input_cols_2[2]:
            price_per_ton = st.number_input("Price per ton", min_value=0.0, step=10.0, format="%.2f")

        current_data = {
            "Project Name": st.session_state.active_project,
            "BOQ Article": st.session_state.loaded_boq,
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

        add_cols = st.columns(2)
        with add_cols[0]:
            if st.button("Add Row"):
                st.session_state.rows.append(current_data.copy())
                st.success("Row added.")
        with add_cols[1]:
            if st.button("Clear All Rows"):
                st.session_state.rows = []
                st.success("All rows cleared.")

    with model_tabs[1]:
        st.subheader("Supplier Input")
        supplier_df = load_supplier_data(st.session_state.active_project)

        s1, s2, s3 = st.columns(3)
        with s1:
            supplier_name_input = st.text_input("Supplier")
        with s2:
            selected_profile_type = st.selectbox("Profile Type", profile_type_options)
        with s3:
            fabric_standard_length_input = st.number_input("Fabric Standard Length", min_value=0.0, step=0.5)

        if st.button("Add Supplier"):
            if supplier_name_input.strip() == "":
                st.warning("Please enter supplier name.")
            else:
                new_row = pd.DataFrame([{
                    "Supplier": supplier_name_input.strip(),
                    "Profile Type": selected_profile_type,
                    "Fabric Standard Length": fabric_standard_length_input
                }])

                supplier_df = supplier_df[
                    ~(
                        (supplier_df["Supplier"].astype(str).str.strip() == supplier_name_input.strip()) &
                        (supplier_df["Profile Type"].astype(str).str.strip() == selected_profile_type)
                    )
                ]

                supplier_df = pd.concat([supplier_df, new_row], ignore_index=True)
                save_supplier_data(supplier_df, st.session_state.active_project)
                st.success("Supplier data saved.")

        supplier_df = load_supplier_data(st.session_state.active_project)

        if st.session_state.edit_mode:
            edited_supplier_df = st.data_editor(
                supplier_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key="supplier_editor"
            )
            st.session_state.supplier_edited_df = edited_supplier_df.copy()
        else:
            st.dataframe(supplier_df, use_container_width=True, hide_index=True)

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

        recalculated_rows = []
        for _, row_item in detail_df.iterrows():
            row_dict = row_item.to_dict()
            row_dict = calculate_row(row_dict, df)
            recalculated_rows.append(row_dict)

        recalculated_df = pd.DataFrame(recalculated_rows).fillna(0)

        summary_df = recalculated_df.groupby(
            ["Floor Level", "Sub Article"], as_index=False
        )[["Number", "Total Treatment Area", "Net Weight", "Weight Incl. Waste", "Total ZBSL", "Total Levering Price"]].sum()

        profile_summary_df = recalculated_df.groupby("Profile", as_index=False).agg({
            "Length": "sum",
            "Number": "sum",
            "Weight Incl. Waste": "sum"
        }).rename(columns={
            "Length": "Total Length",
            "Number": "Total Number",
            "Weight Incl. Waste": "Total Weight"
        })

        supplier_df = load_supplier_data(st.session_state.active_project)
        supplier_options = sorted(supplier_df["Supplier"].dropna().astype(str).str.strip().unique().tolist())

        if supplier_options:
            selected_supplier = supplier_options[0]
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
        else:
            waste_df = pd.DataFrame(columns=["Profile", "Fabric Standard Length", "Supplier Qty", "Waste Length", "Waste Weight"])

        with model_tabs[2]:
            st.subheader("Detail Results")
            if st.session_state.edit_mode:
                edited_df = st.data_editor(
                    recalculated_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key="detail_editor"
                )
                st.session_state.detail_edited_df = edited_df.copy()
            else:
                st.dataframe(recalculated_df, use_container_width=True, hide_index=True)

        with model_tabs[3]:
            st.subheader("Summary by Floor")
            if st.session_state.edit_mode:
                edited_summary_df = st.data_editor(
                    summary_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key="summary_editor"
                )
                st.session_state.summary_edited_df = edited_summary_df.copy()
            else:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

        with model_tabs[4]:
            st.subheader("Profile Sum")
            if st.session_state.edit_mode:
                edited_profile_df = st.data_editor(
                    profile_summary_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key="profile_editor"
                )
                st.session_state.profile_edited_df = edited_profile_df.copy()
            else:
                st.dataframe(profile_summary_df, use_container_width=True, hide_index=True)

        with model_tabs[5]:
            st.subheader("Waste Calculation")
            if st.session_state.edit_mode:
                edited_waste_df = st.data_editor(
                    waste_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key="waste_editor"
                )
                st.session_state.waste_edited_df = edited_waste_df.copy()
            else:
                st.dataframe(waste_df, use_container_width=True, hide_index=True)

    else:
        with model_tabs[2]:
            st.info("No detail rows yet.")
        with model_tabs[3]:
            st.info("No summary yet.")
        with model_tabs[4]:
            st.info("No profile sum yet.")
        with model_tabs[5]:
            st.info("No waste calculation yet.")

if st.session_state.menu_main == "Edit":
    st.subheader("Edit")

    edit_cols = st.columns(3)

    with edit_cols[0]:
        if st.button("Enable Edit Mode", use_container_width=True):
            st.session_state.edit_mode = True
            st.success("All tables are now editable.")

    with edit_cols[1]:
        if st.button("Disable Edit Mode", use_container_width=True):
            st.session_state.edit_mode = False
            st.success("Edit mode off.")

    with edit_cols[2]:
        if st.button("Save All Edited Data", use_container_width=True):
            if "detail_edited_df" in st.session_state:
                edited_df = st.session_state.detail_edited_df.copy()
                recalculated_rows = []
                for _, row_item in edited_df.iterrows():
                    row_dict = row_item.to_dict()
                    row_dict["Project Name"] = st.session_state.active_project
                    row_dict["BOQ Article"] = st.session_state.loaded_boq
                    row_dict = calculate_row(row_dict, df)
                    recalculated_rows.append(row_dict)
                st.session_state.rows = recalculated_rows
                save_results(st.session_state.rows, st.session_state.active_project)

            if "supplier_edited_df" in st.session_state:
                save_supplier_data(st.session_state.supplier_edited_df.copy(), st.session_state.active_project)

            st.success(f"All edited data saved to project: {st.session_state.active_project}")

    st.info("Click Enable Edit Mode, then go to Model and edit tables. After editing, come back here and click Save All Edited Data.")

if st.session_state.menu_main == "Calculation":
    st.subheader("Calculation")
    st.info("This section is ready for future calculation code.")

    calc_option = st.selectbox(
        "Future Calculation Menu",
        [
            "Select",
            "Connection Design",
            "Bolt Calculation",
            "Weld Calculation",
            "Plate Check",
            "Optimization",
            "Custom Code"
        ]
    )

    if calc_option != "Select":
        st.write(f"Selected: {calc_option}")
        st.write("You can add your next calculation code here.")
