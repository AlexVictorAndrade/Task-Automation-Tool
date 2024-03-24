Task Automation Tool Documentation

    Overview
        This task automation tool is designed to streamline repetitive tasks by automating the process of copying data from an Excel spreadsheet and pasting it into a web form. It operates in a loop, processing each row of the Excel spreadsheet until it reaches the last row.

    Installation
        Ensure you have the following libraries installed:

        openpyxl
        pyperclip
        pyautogui
        You can install them using pip: "pip install openpyxl pyperclip pyautogui"

    Usage
        1. Prepare your Excel spreadsheet with the data you want to copy into the web form. Make sure the data is organized in columns corresponding to the fields in the web form.

        2. Update the file path in the script to point to your Excel spreadsheet:
            "workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')"

        3. Run the script. Click on the play button.

    Functionality
        1. Copying and Pasting: 
            The script iterates through each row of the Excel spreadsheet and copies data from specific columns. It then pastes this data into corresponding fields in the web form.

        2. Button Clicks: 
            It simulates mouse clicks to navigate through the web form, such as clicking on the "Next" button to move to the next set of fields or the "Finish" button to complete the form.

        3. Handling Dropdowns and Radio Buttons: 
            It handles dropdown menus and radio buttons based on predefined conditions. For example, selecting options from dropdown menus based on the data provided in the Excel spreadsheet.

    Customization
        1. Field Coordinates: 
            The script uses hardcoded screen coordinates for clicking on form fields. You may need to adjust these coordinates based on your screen resolution and the layout of the web form.

        2. Dropdown and Radio Button Logic: 
            If your web form has different options or logic for dropdown menus and radio buttons, you will need to modify the script to accommodate these variations.

    Limitations
        1. Screen Resolution: The script's reliance on screen coordinates makes it sensitive to changes in screen resolution or layout of the web form.


    Disclaimer
        1. This tool is provided as-is without any guarantees. Use it responsibly and ensure it complies with any applicable terms of service or usage agreements for the web form you are automating.

    Author
        This script was developed by Alex Victor de Andrade.
        andrade.alexvictor@gmail.com

    License
        This project is licensed under the MIT License.