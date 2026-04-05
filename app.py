import streamlit as st

st.title("Steel Calculation Tool")

profile = st.text_input("Profile (e.g. HEA200)")
length = st.number_input("Length (m)", min_value=0.0)
number = st.number_input("Number of elements", min_value=1)

if st.button("Calculate"):
    total_length = length * number
    st.write("Total length:", total_length, "m")
