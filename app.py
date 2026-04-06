import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="SEST", layout="wide")
st.title("SEST")

# ---------- STYLE (SMALL HORIZONTAL MENU) ----------
st.markdown("""
<style>
div.stButton > button {
    height: 1.9rem;
    min-height: 1.9rem;
    padding: 0rem 0.5rem;
    font-size: 13px;
    border-radius: 6px;
}
</style>
""", unsafe_allow_html=True)

# ---------- SESSION ----------
if "menu_main" not in st.session_state:
    st.session_state.menu_main = "Model"

if "edit_mode" not in st.session_state:
    st.session_state.edit_mode = False

if "rows" not in st.session_state:
    st.session_state.rows = []

if "project_name" not in st.session_state:
    st.session_state.project_name = "default_project"

PROJECT_DIR = "projects"
os.makedirs(PROJECT_DIR, exist_ok=True)

# ---------- TOOLBAR ----------
cols = st.columns([0.6, 0.8, 0.6, 1.0, 6])

with cols[0]:
    if st.button("File"):
        st.session_state.menu_main = "File"

with cols[1]:
    if st.button("Model"):
        st.session_state.menu_main = "Model"

with cols[2]:
    if st.button("Edit"):
        st.session_state.menu_main = "Edit"

with cols[3]:
    if st.button("Calculation"):
        st.session_state.menu_main = "Calculation"

st.divider()

# ---------- FILE ----------
if st.session_state.menu_main == "File":
    st.subheader("File")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if st.button("New Project"):
            st.session_state.rows = []
            st.session_state.project_name = "default_project"
            st.success("New project created")

    with c2:
        project_name = st.text_input("Project Name", value=st.session_state.project_name)

    with c3:
        if st.button("Save Project"):
            df = pd.DataFrame(st.session_state.rows)
            df.to_excel(f"{PROJECT_DIR}/{project_name}.xlsx", index=False)
            st.session_state.project_name = project_name
            st.success("Saved")

    with c4:
        files = [f.replace(".xlsx","") for f in os.listdir(PROJECT_DIR)]
        selected = st.selectbox("Open Project", files if files else ["default_project"])

        if st.button("Load"):
            try:
                df = pd.read_excel(f"{PROJECT_DIR}/{selected}.xlsx")
                st.session_state.rows = df.to_dict("records")
                st.session_state.project_name = selected
                st.success("Loaded")
            except:
                st.warning("No file found")

# ---------- MODEL ----------
if st.session_state.menu_main == "Model":
    st.subheader("Model")

    tabs = st.tabs([
        "Data Input",
        "Supplier Input",
        "Detail Results",
        "Summary by Floor"
    ])

    # ---------- DATA INPUT ----------
    with tabs[0]:
        c1, c2, c3 = st.columns(3)

        with c1:
            floor = st.selectbox("Floor", ["G","1","2","3","4"])

        with c2:
            profile = st.text_input("Profile")

        with c3:
            length = st.number_input("Length", 0.0)

        c4, c5 = st.columns(2)

        with c4:
            qty = st.number_input("Quantity", 1)

        with c5:
            price = st.number_input("Price/t", 0.0)

        if st.button("Add"):
            st.session_state.rows.append({
                "Floor": floor,
                "Profile": profile,
                "Length": length,
                "Quantity": qty,
                "Price/t": price
            })
            st.success("Added")

    # ---------- SUPPLIER ----------
    with tabs[1]:
        st.info("Supplier table")

        supplier_df = pd.DataFrame({
            "Supplier": ["A","B"],
            "Type": ["I","RHS"],
            "Length": [12,6]
        })

        if st.session_state.edit_mode:
            st.data_editor(supplier_df)
        else:
            st.dataframe(supplier_df)

    # ---------- DETAIL ----------
    with tabs[2]:
        if st.session_state.rows:
            df = pd.DataFrame(st.session_state.rows)

            if st.session_state.edit_mode:
                edited = st.data_editor(df)
                st.session_state.rows = edited.to_dict("records")
            else:
                st.dataframe(df)
        else:
            st.info("No data")

    # ---------- SUMMARY ----------
    with tabs[3]:
        if st.session_state.rows:
            df = pd.DataFrame(st.session_state.rows)
            summary = df.groupby("Floor")["Quantity"].sum().reset_index()

            if st.session_state.edit_mode:
                st.data_editor(summary)
            else:
                st.dataframe(summary)
        else:
            st.info("No summary yet")

# ---------- EDIT ----------
if st.session_state.menu_main == "Edit":
    st.subheader("Edit")

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Enable Edit"):
            st.session_state.edit_mode = True
            st.success("Edit mode ON")

    with c2:
        if st.button("Disable Edit"):
            st.session_state.edit_mode = False
            st.success("Edit mode OFF")

    if st.button("Save Changes"):
        df = pd.DataFrame(st.session_state.rows)
        df.to_excel(f"{PROJECT_DIR}/{st.session_state.project_name}.xlsx", index=False)
        st.success("Saved")

# ---------- CALCULATION ----------
if st.session_state.menu_main == "Calculation":
    st.subheader("Calculation")

    st.selectbox("Future tools", [
        "Connection",
        "Bolt",
        "Weld",
        "Plate"
    ])

    st.info("Add your future calculation code here")
