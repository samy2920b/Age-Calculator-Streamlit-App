import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Function to calculate age from DOB
def calculate_age(dob):
    # Current date
    today = datetime.today()
   
    # Check if dob is not NaT
    if pd.isna(dob):
        return None
   
    # Calculate difference in years, months, and days
    age_years = today.year - dob.year
    age_months = today.month - dob.month
    age_days = today.day - dob.day
   
    # Adjust for negative months or days
    if age_months < 0:
        age_years -= 1
        age_months += 12
   
    if age_days < 0:
        age_months -= 1
        age_days += (datetime(today.year, today.month, 1) - datetime(today.year, today.month-1, 1)).days
   
    # Return formatted age
    return f"{age_years} years, {age_months} months, {age_days} days"

# Streamlit app
def main():
    st.title("Age Calculator from Excel")
    st.subheader('Made by Saksham Srivastava')
   
    # Option to download an Excel template
    if st.button("Download Template"):
        template = pd.DataFrame(columns=["Name", "DOB"])
        output = BytesIO()
        template.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        st.download_button(
            label="Download Excel Template",
            data=output,
            file_name="template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
   
    # Step 1: Upload Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
   
    if uploaded_file:
        # Step 2: Read the uploaded Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)
       
        # Step 3: Check if 'DOB' column exists
        if 'DOB' not in df.columns or 'Name' not in df.columns:
            st.error("The Excel file must contain 'Name' and 'DOB' columns.")
            return
       
        # Step 4: Convert DOB column to datetime
        df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce')
       
        # Step 5: Calculate age for each row
        df['Age'] = df['DOB'].apply(calculate_age)
       
        # Step 6: Display the updated DataFrame with calculated ages
        st.write("Updated Data with Age Calculation:")
        st.dataframe(df)
       
        # Step 7: Option to download the updated file
        if st.button('Download the updated Excel file'):
            # Save the DataFrame to an Excel file
            output = BytesIO()
            df.to_excel(output, index=False, engine='xlsxwriter')
            output.seek(0)
            st.download_button(
                label="Download Age Calculated Excel File",
                data=output,
                file_name="age_calculated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
