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
def extract_attendance(file_obj, start_date, end_date):
    attendance_data = []
    time_in = None
    pattern = r'\[(.*?)\] (.*?):\s*(.*?\b(?:clock(?:ed\s*out|ing\s*in|ing\s*out)|morning[,\s]+clock(?:ing\s*in|ing\s*out))\b.*?)\s*'
    pattern_in = r'\b(clock|in)\b'
    pattern_out = r'\b(clock|out)\b'

    for line in file_obj:
        match = re.search(pattern, line.lower(), re.IGNORECASE)
        if match:
            groups = match.groups()
            if len(groups) == 3:
                date_str, username, status = groups
                date = datetime.strptime(date_str, '%d/%m/%Y, %H:%M:%S')
                if start_date <= date.date() <= end_date:
                    matches_in = re.findall(pattern_in, status, re.IGNORECASE)
                    matches_out = re.findall(pattern_out, status, re.IGNORECASE)
                    if matches_in:
                        time_in = date.time()
                        attendance_data.append({
                            'Date': date.date(),
                            'Time_In': time_in.strftime('%H:%M'),
                            'Time_Out': 'Holiday',
                            'Username': username.strip()
                        })
                    elif matches_out:
                        if time_in:
                            attendance_data.append({
                                'Date': date.date(),
                                'Time_In': time_in.strftime('%H:%M'),
                                'Time_Out': date.time().strftime('%H:%M'),
                                'Username': username.strip()
                            })
                            time_in = None

    attendance_df = pd.DataFrame(attendance_data)

    if 'Date' in attendance_df.columns:
        attendance_df['Date'] = pd.to_datetime(attendance_df['Date'])

    return attendance_df

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
        # attendance_df = extract_attendance(file_stream, start_date, end_date, whatsapp_name)
        attendance_df = extract_attendance(file_stream, start_date, end_date)


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