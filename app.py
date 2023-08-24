
import streamlit as st
import pandas as pd
import openpyxl as px
import openai

# OpenAI API Key (Replace with your actual key)
openai_api_key = "YOUR_OPENAI_API_KEY"

# Function to analyze data with OpenAI
def analyze_data_with_openai(data):
    data_text = data.to_string()
    prompt = f"Analyze the following business data and provide insights, budget analysis, and forecasting:\n{data_text}"
    response = openai.Completion.create(engine="text-davinci-003", prompt=prompt, max_tokens=150)
    analysis_result = response.choices[0].text
    return analysis_result

# Function to review Excel file
def review_excel(file):
    workbook = px.load_workbook(file)
    sheet = workbook.active
    data = pd.DataFrame(sheet.values)
    data.columns = data.iloc[0]
    data = data[1:]
    analysis_result = analyze_data_with_openai(data)
    return analysis_result

# Function to check formulas and provide feedback
def check_formulas(file):
    workbook = px.load_workbook(file)
    sheet = workbook.active
    feedback = []
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                feedback.append(f"Formula in cell {cell.coordinate}: {cell.value}")
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("#"):
                feedback.append(f"Error in cell {cell.coordinate}: {cell.value}")
    return feedback

# Main app
st.title("Excel Business Analysis Reviewer")
st.sidebar.title("Navigation")
page = st.sidebar.selectbox("Choose a page", ["Home", "Ask Questions"])

if page == "Home":
    st.subheader("Upload Excel File for Analysis")
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    if uploaded_file:
        st.write("Reviewing the Excel file...")
        analysis_result = review_excel(uploaded_file)
        st.subheader("Analysis Results:")
        st.write(analysis_result)

        st.subheader("Formula Check:")
        formula_feedback = check_formulas(uploaded_file)
        for item in formula_feedback:
            st.write(item)

elif page == "Ask Questions":
    st.subheader("Ask Questions About Excel")
    user_question = st.text_input("Enter your question:")
    if user_question:
        st.write("Thank you for your question! We will respond as soon as possible.")

# Custom styling
st.markdown("""
<style>
body {
    color: #4f8bf9;
    background-color: #f0f0f5;
    font-family: Arial, sans-serif;
}
</style>
""", unsafe_allow_html=True)
