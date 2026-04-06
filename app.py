import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="SEST", layout="wide")
st.title("SEST")

if "edit_mode" not in st.session_state:
    st.session_state.edit_mode = False

if "rows" not in st.session_state:
    st.session_state.rows = []

if "project_name" not in st.session_state:
    st.session_state.project_name = "default_project"

PROJECT_DIR = "projects"
os.makedirs(PROJECT_DIR, exist_ok=True)

main_tabs = st.tabs(["File", "Model", "Edit", "Calculation"])

with main_tabs[0]:
    st.subheader("File")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if st.button("New Project"):
            st.session_state.rows = []
            st.session_state.project_name = "default_project"
            st.success("New project created")

    with c2:
        project_name = st.text_input("Project Name", value=st.session_state.project_name, key="file_project_name")

    with c3:
        if st.button("Save Project"):
            df = pd.DataFrame(st.session_state.rows)
            df.to_excel(f"{PROJECT_DIR}/{project_name}.xlsx", index=False)
            st.session_state.project_name = project_name
            st.success("Saved")

    with c4:
        files = [f.replace(".xlsx", "") for f in os.listdir(PROJECT_DIR) if f.endswith(".xlsx")]
        selected = st.selectbox("Open Project", files if files else ["default_project"], key="open_project")

        if st.button("Load"):
            try:
                df = pd.read_excel(f"{PROJECT_DIR}/{selected}.xlsx")
                st.session_state.rows = df.to_dict("records")
                st.session_state.project_name = selected
                st.success("Loaded")
            except:
                st.warning("No file found")

with main_tabs[1]:
    st.subheader("Model")

    model_tabs = st.tabs([
        "Data Input",
        "Supplier Input",
        "Detail Results",
        "Summary by Floor"
    ])

    with model_tabs[0]:
        st.text_input("Project Name", value=st.session_state.project_name, disabled=True)

        boq = st.text_input("BOQ Article", key="boq_article")

        c1, c2, c3 = st.columns(3)
        with c1:
            floor = st.selectbox("Floor", ["G", "1", "2", "3", "4"])
        with c2:
            profile = st.text_input("Profile")
        with c3:
            length = st.number_input("Length", min_value=0.0, value=0.0)

        c4, c5 = st.columns(2)
        with c4:
            qty = st.number_input("Quantity", min_value=1, value=1)
        with c5:
            price = st.number_input("Price/t", min_value=0.0, value=0.0)

        c6, c7 = st.columns(2)
        with c6:
            if st.button("Add Row"):
                st.session_state.rows.append({
                    "Project Name": st.session_state.project_name,
                    "BOQ Article": boq,
                    "Floor": floor,
                    "Profile": profile,
                    "Length": length,
                    "Quantity": qty,
                    "Price/t": price
                })
                st.success("Added")

        with c7:
            if st.button("Clear Rows"):
                st.session_state.rows = []
                st.success("Cleared")

    with model_tabs[1]:
        st.subheader("Supplier Input")

        supplier_df = pd.DataFrame({
            "Supplier": ["A", "B"],
            "Type": ["I", "RHS"],
            "Length": [12, 6]
        })

        if st.session_state.edit_mode:
            st.data_editor(supplier_df, use_container_width=True, key="supplier_editor")
        else:
            st.dataframe(supplier_df, use_container_width=True, hide_index=True)

    with model_tabs[2]:
        st.subheader("Detail Results")

        if st.session_state.rows:
            df = pd.DataFrame(st.session_state.rows)

            if st.session_state.edit_mode:
                edited = st.data_editor(df, use_container_width=True, key="detail_editor")
                st.session_state.rows = edited.to_dict("records")
            else:
                st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("No data")

    with model_tabs[3]:
        st.subheader("Summary by Floor")

        if st.session_state.rows:
            df = pd.DataFrame(st.session_state.rows)
            if "Floor" in df.columns and "Quantity" in df.columns:
                summary = df.groupby("Floor", as_index=False)["Quantity"].sum()

                if st.session_state.edit_mode:
                    st.data_editor(summary, use_container_width=True, key="summary_editor")
                else:
                    st.dataframe(summary, use_container_width=True, hide_index=True)
            else:
                st.info("No summary yet")
        else:
            st.info("No summary yet")

with main_tabs[2]:
    st.subheader("Edit")

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("Enable Edit"):
            st.session_state.edit_mode = True
            st.success("Edit mode ON")

    with c2:
        if st.button("Disable Edit"):
            st.session_state.edit_mode = False
            st.success("Edit mode OFF")

    with c3:
        if st.button("Save Changes"):
            df = pd.DataFrame(st.session_state.rows)
            df.to_excel(f"{PROJECT_DIR}/{st.session_state.project_name}.xlsx", index=False)
            st.success("Saved")

with main_tabs[3]:
    st.subheader("Calculation")

    st.selectbox("Future tools", [
        "Connection",
        "Bolt",
        "Weld",
        "Plate"
    ])

    st.info("Add your future calculation code here")
