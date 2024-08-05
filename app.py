import attendance
import streamlit as st 
from io import BytesIO

def main():
    st.title("Attendance Management System")
    st.write("This is a simple web app to manage attendance of employees")
    st.write("Please upload the attendance file")
    attendance_file = st.file_uploader("Choose a file", key="attendance_file")

    st.write("Please upload the user data file")
    user_data_file = st.file_uploader("Choose a file", key="user_data_file")

    if st.button("Submit"):
        if attendance_file is None or user_data_file is None:
            st.write("Please upload the files")
            return
        
        attendance_obj = attendance.Attendance(attendance_file, user_data_file)

        try:
            attendance_obj.parse_input()

        except ValueError:
            st.write("Invalid attendance file")
            return

        attendance_obj.parse_entries()

        try:
            attendance_obj.map_employees()
            
        except ValueError:
            st.write("Invalid user data file")
            return
        
        workbook, file_name = attendance_obj.write_excel()

        output = BytesIO()
        workbook.save(output)
        output.seek(0)

        # Provide a download button
        st.download_button(
            label="Download Attendance Workbook",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()