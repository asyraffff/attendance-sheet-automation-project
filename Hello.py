import streamlit as st
import pandas as pd
import re
from datetime import datetime
from docx import Document
from io import StringIO
from streamlit.logger import get_logger

LOGGER = get_logger(__name__)

# Function to extract attendance data
# @st.cache_data
def extract_attendance(file_obj, start_date, end_date, whatsapp_name):
    attendance_data = []
    time_in = None
    # pattern = r'\[(.*?)\] {}:\s*(.*?\b(?:clock(?:ed\s*out|ing\s*in|ing\s*out)|morning[,\s]+clock(?:ing\s*in|ing\s*out))\b.*?)\s*'.format(whatsapp_name.lower())
    pattern = r'\[(.*?)\] ((?:\+\d{1,3}\s?\d{1,2}\u200E?-\d{3}\s?\d{4})|(?:\w+(?:\s\w+)*(?:\s\w+\s\w+)*)?):\s*(.*?\b(?:clock(?:ed\s*out|ing\s*in|ing\s*out)|morning[,\s]+clock(?:ing\s*in|ing\s*out))\b.*?)\s*'
    pattern_in = r'\b(clock|in)\b'
    pattern_out = r'\b(clock|out)\b'

    # with open(file_path, 'r', encoding='utf-8') as f:
    for line in file_obj:
        match = re.search(pattern, line.lower(), re.IGNORECASE)
        if match:
            date_str, status = match.groups()
            date = datetime.strptime(date_str, '%d/%m/%Y, %H:%M:%S')
            if start_date <= date.date() <= end_date:
                matches_in = re.findall(pattern_in, status, re.IGNORECASE)
                matches_out = re.findall(pattern_out, status, re.IGNORECASE)
                if matches_in:
                    time_in = date.time()
                elif matches_out:
                    if time_in:
                        attendance_data.append({
                            'Date': date.date(),
                            'Time_In': time_in.strftime('%H:%M'),
                            'Time_Out': date.time().strftime('%H:%M')
                        })
                        time_in = None

    # Create a DataFrame from the attendance data
    attendance_df = pd.DataFrame(attendance_data)
    # attendance_df = pd.DataFrame(attendance_data, columns=['Date', 'Time_In', 'Time_Out'])

    # Convert the 'Date' column to datetime64[ns] data type
    attendance_df['Date'] = pd.to_datetime(attendance_df['Date'])

    # Create a range of dates between start_date and end_date
    date_range = pd.date_range(start_date, end_date, freq='D')

    # Create a DataFrame with the date range
    date_range_df = pd.DataFrame({'Date': date_range})

    # Merge the extracted attendance data with the date range DataFrame
    merged_df = date_range_df.merge(attendance_df, how='left', on='Date')

    # Remove the time component from the Date column
    merged_df['Date'] = merged_df['Date'].dt.date

    # Fill missing values with 'Holiday'
    merged_df = merged_df.fillna({'Time_In': 'Holiday', 'Time_Out': 'Holiday'})

    return merged_df

# Streamlit app
def app():
    st.set_page_config(
        page_title="Attendance Sheet Automation System (ASAS)",
        page_icon="👍",
    )

    st.title("Attendance Sheet Automation System (ASAS)")

    # File upload
    file = st.file_uploader("Upload WhatsApp chat file", type="txt")

    # Input fields
    start_date = st.date_input("Start Date", value=datetime(2024, 3, 1).date())
    end_date = st.date_input("End Date", value=datetime(2024, 3, 31).date())
    whatsapp_name = st.text_input("WhatsApp Name", value="asyraf")

    if file:
        # Read the contents of the uploaded file
        file_content = file.getvalue().decode("utf-8")

        # Create a StringIO object with the file content
        file_stream = StringIO(file_content)

        # Extract attendance data
        attendance_df = extract_attendance(file_stream, start_date, end_date, whatsapp_name)

        # Display attendance data
        st.write("Attendance Data:")
        st.dataframe(attendance_df)

        st.write("Trainee Details:")
        # Get user input
        name = st.text_input("Name", value="Amirul Asyraf")
        dept_div = st.text_input("Department/Division", value="GTS-Energy")
        phone_number = st.text_input("Phone Number", value="O13-35097092")
        email = st.text_input("Email", value="amirulasyraf.zainala@petronas.com")
        ic_number = st.text_input("IC Number", value="011102-14-0551")
        cost_centre = st.text_input("Cost - Centre", value="001PCT-606")
        ext_no = st.text_input("EXT NO")
        contact_person = st.text_input("Contact Person", value="Rafidah Bt. Ramli")

        # start_date = st.text_input("Start Date")
        # end_date = st.text_input("End Date")
        # whatsapp_name = st.text_input("Whatsapp name")

        position = st.text_input("Position", value="Intern")
        # date_today = st.date_input("Date")
        date_today = st.text_input("Date Today", value="2024/03/01")
        sv_name = st.text_input("Name (Supervisor)", value="Tanasegaran Kuppusamy")
        sv_position = st.text_input("Position (Supervisor)", value="Staff (Energy)")
        # sv_date = st.text_input("Date (Supervisor)")

        # Export to Word document
        if st.button("Export to Word"):
            doc = Document('Attendance_Sheet.docx')
            table = doc.tables[0]

            # Fill in personal details
            table.cell(2, 2).text = name
            table.cell(2, 9).text = ic_number
            table.cell(3, 2).text = dept_div
            table.cell(3, 9).text = cost_centre
            table.cell(4, 2).text = phone_number
            table.cell(4, 9).text = ext_no
            table.cell(5, 2).text = email
            table.cell(5, 10).text = contact_person
            table.cell(44, 2).text = name
            table.cell(44, 9).text = sv_name
            table.cell(45, 2).text = position
            table.cell(45, 9).text = sv_position
            table.cell(46, 2).text = date_today

            # Write attendance data into the table
            for i, row in attendance_df.iterrows():
                # print(f"i : {i} , row : {row}")
                for j, value in enumerate(row):
                    # print(f"j : {j} , value : {value}")
                    table.cell(i + 10, j + 3).text = str(value)

            # Save the modified document
            modified_doc_path = 'modified_document.docx'
            doc.save(modified_doc_path)
            st.success("Attendance data exported to Word document successfully!")

            # Provide download link for the modified document
            with open(modified_doc_path, "rb") as file:
                btn = st.download_button(
                    label="Download Modified Document",
                    data=file,
                    file_name="modified_document.docx",
                    mime="application/octet-stream"
                )
    else:
        st.warning("Please upload a WhatsApp chat file.")

if __name__ == "__main__":
    app()