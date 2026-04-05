import streamlit as st
import pandas as pd
import io
import os
import math

st.set_page_config(page_title="Steel Calculation App", layout="wide")
st.title("Steel Calculation App")

RESULTS_FILE = "results.xlsx"
PROFILES_FILE = "Profiles.xlsx"
MAX_PIECE_LENGTH = 23.0

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

    if profile_name.startswith("HE"):
        return "HE"
    elif profile_name.startswith("K"):
        return "RHS"
    elif profile_name.startswith("R"):
        return "CHS"
    elif profile_name.startswith("L"):
        return "L"
    else:
        return "Other"

def get_standard_length(profile_name, he_len, rhs_len, chs_len, l_len):
    profile_type = get_profile_type(profile_name)

    if profile_type == "HE":
        return he_len
    elif profile_type == "RHS":
        return rhs_len
    elif profile_type == "CHS":
        return chs_len
    elif profile_type == "L":
        return l_len
    else:
        return 0.0

def calculate_row(row_data, profile_df):
    profile_name = str(row_data.get("Profile", "")).strip()

    default_result = {
        "Split Pieces": 1,
        "kg/m": 0.0,
        "Total Treatment Area": 0.0,
        "Total Weight": 0.0,
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

    total_weight = kgm * calc_length * calc_number
    total_treatment_area = m2_per_m * calc_length * calc_number
    total_zbsl = zbsl * calc_number
    total_price = (total_weight / 1000) * price_per_ton

    row_data["Length"] = round(calc_length, 2)
    row_data["Number"] = int(calc_number) if float(calc_number).is_integer() else round(calc_number, 2)
    row_data["Split Pieces"] = int(split_pieces)
    row_data["Price/t"] = round(price_per_ton, 2)
    row_data["kg/m"] = round(kgm, 2)
    row_data["Total Treatment Area"] = round(total_treatment_area, 2)
    row_data["Total Weight"] = round(total_weight, 2)
    row_data["Total ZBSL"] = round(total_zbsl, 2)
    row_data["Total Levering Price"] = round(total_price, 2)

    return row_data

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

def save_results(rows):
    pd.DataFrame(rows).to_excel(RESULTS_FILE, index=False)

def load_saved_results():
    if os.path.exists(RESULTS_FILE):
        saved_df = pd.read_excel(RESULTS_FILE).fillna("")
        return saved_df.to_dict("records")
    return []

df = load_profiles()
profile_list = df["Profile"].dropna().astype(str).str.strip().tolist()

floor_options = ["Ground Floor", "First Floor", "Second Floor", "Third Floor", "Fourth Floor"]
sub_article_options = ["Beam", "Column", "Brace", "Plate", "Connection"]

if "rows" not in st.session_state:
    st.session_state.rows = load_saved_results()

st.subheader("Project Information")
col1, col2 = st.columns(2)

with col1:
    project = st.text_input("Project Name")

with col2:
    boq = st.text_input("BOQ Article")

st.subheader("Fabric Standard Lengths")
f1, f2, f3, f4 = st.columns(4)

with f1:
    he_standard_length = st.number_input("HE Standard Length (m)", min_value=0.0, value=12.0, step=0.5)

with f2:
    rhs_standard_length = st.number_input("RHS Standard Length (m), K...", min_value=0.0, value=6.0, step=0.5)

with f3:
    chs_standard_length = st.number_input("CHS Standard Length (m), R...", min_value=0.0, value=6.0, step=0.5)

with f4:
    l_standard_length = st.number_input("L Standard Length (m), L...", min_value=0.0, value=6.0, step=0.5)

st.subheader("Input Data")
col3, col4, col5, col6, col7, col8 = st.columns(6)

with col3:
    floor_level = st.selectbox("Floor Level", floor_options)

with col4:
    sub_article = st.selectbox("Sub Article", sub_article_options)

with col5:
    profile = st.selectbox("Profile", profile_list)

with col6:
    length = st.number_input("Length (m)", min_value=0.0, step=0.1, format="%.2f")

with col7:
    quantity = st.number_input("Quantity", min_value=1, step=1)

with col8:
    price_per_ton = st.number_input("Price per ton", min_value=0.0, step=10.0, format="%.2f")

current_data = {
    "Project Name": project,
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
    "Total Weight": 0.0,
    "Total ZBSL": 0.0,
    "Total Levering Price": 0.0
}

current_data = calculate_row(current_data, df)

st.subheader("Current Result")
current_result_df = pd.DataFrame([current_data])[[
    "Floor Level", "Sub Article", "Profile", "Length", "Number", "Price/t"
]]
st.dataframe(current_result_df, use_container_width=True, hide_index=True)

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("Add"):
        st.session_state.rows.append(current_data.copy())
        save_results(st.session_state.rows)
        st.success("Row added and saved.")

with col_btn2:
    if st.button("Clear All Rows"):
        st.session_state.rows = []
        if os.path.exists(RESULTS_FILE):
            os.remove(RESULTS_FILE)
        st.success("All rows cleared.")

st.subheader("Detail Results")

if st.session_state.rows:
    detail_df = pd.DataFrame(st.session_state.rows)

    expected_columns = [
        "Project Name", "BOQ Article", "Floor Level", "Sub Article", "Profile",
        "Length", "Number", "Price/t", "Split Pieces", "kg/m",
        "Total Treatment Area", "Total Weight", "Total ZBSL", "Total Levering Price"
    ]

    for col in expected_columns:
        if col not in detail_df.columns:
            detail_df[col] = ""

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
            "Total Weight": st.column_config.NumberColumn("Total Weight", disabled=True, format="%.2f"),
            "Total ZBSL": st.column_config.NumberColumn("Total ZBSL", disabled=True, format="%.2f"),
            "Total Levering Price": st.column_config.NumberColumn("Total Levering Price", disabled=True, format="%.2f"),
        }
    )

    recalculated_rows = []
    for _, row_item in edited_df.iterrows():
        row_dict = row_item.to_dict()
        row_dict = calculate_row(row_dict, df)
        recalculated_rows.append(row_dict)

    recalculated_df = pd.DataFrame(recalculated_rows)
    st.session_state.rows = recalculated_df.to_dict("records")
    save_results(st.session_state.rows)

    st.dataframe(recalculated_df, use_container_width=True, hide_index=True)

    st.subheader("Summary by Floor Level and Sub Article")
    summary_df = recalculated_df.groupby(
        ["Floor Level", "Sub Article"], as_index=False
    )[["Number", "Total Treatment Area", "Total Weight", "Total ZBSL", "Total Levering Price"]].sum()
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.subheader("Summary by Profile")
    profile_summary_df = recalculated_df.groupby("Profile", as_index=False).agg({
        "Length": "sum",
        "Number": "sum",
        "kg/m": "first"
    })

    profile_summary_df = profile_summary_df.rename(columns={
        "Length": "Total Length",
        "Number": "Total Number"
    })

    profile_summary_df["Profile Type"] = profile_summary_df["Profile"].apply(get_profile_type)

    profile_summary_df["Standard Length"] = profile_summary_df["Profile"].apply(
        lambda x: get_standard_length(
            x,
            he_standard_length,
            rhs_standard_length,
            chs_standard_length,
            l_standard_length
        )
    )

    profile_summary_df["Required Standard Bars"] = profile_summary_df.apply(
        lambda row: math.ceil(row["Total Length"] / row["Standard Length"])
        if to_float(row["Standard Length"]) > 0 else 0,
        axis=1
    )

    profile_summary_df["Waste Length"] = profile_summary_df.apply(
        lambda row: round(
            row["Required Standard Bars"] * row["Standard Length"] - row["Total Length"], 2
        ) if to_float(row["Standard Length"]) > 0 else 0,
        axis=1
    )

    profile_summary_df["Waste Weight"] = profile_summary_df.apply(
        lambda row: round(row["Waste Length"] * row["kg/m"], 2),
        axis=1
    )

    profile_summary_df = profile_summary_df[
        [
            "Profile",
            "Profile Type",
            "Total Length",
            "Total Number",
            "kg/m",
            "Standard Length",
            "Required Standard Bars",
            "Waste Length",
            "Waste Weight"
        ]
    ]

    st.dataframe(profile_summary_df, use_container_width=True, hide_index=True)

    total_waste_weight = round(profile_summary_df["Waste Weight"].sum(), 2)
    st.metric("Total Waste Weight", total_waste_weight)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        recalculated_df.to_excel(writer, index=False, sheet_name="Detail Results")
        summary_df.to_excel(writer, index=False, sheet_name="Summary by Floor")
        profile_summary_df.to_excel(writer, index=False, sheet_name="Summary by Profile")

        total_waste_df = pd.DataFrame({
            "Item": ["Total Waste Weight"],
            "Value": [total_waste_weight]
        })
        total_waste_df.to_excel(writer, index=False, sheet_name="Total Waste")

    output.seek(0)

    st.download_button(
        label="Export to Excel",
        data=output,
        file_name="steel_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("No rows added yet.")
