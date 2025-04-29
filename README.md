# **Student Registration and Database System**
## Objective
  
The objective of this system is to provide a user-friendly graphical interface (GUI) for collecting student data and storing it in an Excel spreadsheet for easy organization and retrieval.

## Functionality

The GUI application provides functionalities for:

1. **Data Entry:** Students or authorized personnel can enter student information through various input fields, including:

    - **Personal Information:** Full Name, Matriculation Number, Gender (Radio Buttons), Religion (Combobox)

    - **General Information:** Date (automatically populated), College (Combobox), Department, Date of Birth

    - **Contact Information:** Email Address, Phone Number

2. **Image Upload:** Students can upload a profile picture using the "Upload" button. The image is then saved with the student's Matriculation Number as the filename. (Optional: Error handling for file types is not implemented)

3. **Data Validation:** Basic validation checks ensure required fields are filled before saving (Name, Level, Religion, etc.) but more extensive validation might be required.

4. **Data Saving:** Clicking the "Save" button triggers the following actions:

    - **Excel File Management:** The system checks if the "Student_data.xlsx" file exists. If not, it creates a new file with headers for each data point.

    - **Data Writing:** The entered student information is written to a new row in the Excel spreadsheet.

    - **Image Saving (Optional):** If an image is uploaded, it's saved in a separate folder named "Student_images" with the student's Matriculation Number as the filename. (Error handling for saving images is not implemented)

5. **Reset Functionality:** Clicking the "Reset" button clears all data entry fields and the profile picture to its default state, allowing for new student registration.

6. **Exit Functionality:** Clicking the "Exit" button closes the application. 


## Important Information

- The system relies on two files:

    - **Student_data.xlsx:** This Excel file stores all registered student information.

    - **Student_images (folder):** This folder (needs to be created manually) stores the uploaded profile pictures of students (if the image upload functionality is used).

- Error handling is implemented for missing data entries before saving.

- Consider adding more extensive data validation for student information like email format or phone number format.

- The code utilizes Combobox elements for selecting options from predefined lists (Level, Religion, College). You can modify these lists as needed.

## - Using the Application:

    - Fill in the student details in the provided fields.

    - Upload a profile picture using the "Upload" button.

    - Click "Save" to store the data in the Excel file and save the profile picture.

    - Use the "Reset" button to clear the form.

    - Click "Exit" to close the application.

  

## Summary

This Student Registration and Database System provides an easy-to-use interface for collecting and storing student information. The use of Tkinter for the GUI, openpyxl for Excel file handling, and PIL for image processing ensures a seamless and efficient registration process.


![Screenshot 2024-06-07 145530](https://github.com/user-attachments/assets/0bd7e29e-4da9-4881-86a0-74222b58ed6d)
