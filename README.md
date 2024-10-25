# Subcontractor Evaluation System

This project was developed to streamline the scoring and management process for drivers and suppliers based on specific criteria. The system enables users to select and fill information directly from an Excel spreadsheet and post scores on separate tabs for drivers and suppliers.

**KEY FEATURES**

1. **Driver Scoring**  
   The system allows users to enter scores for drivers based on specific criteria. Users can select the driver, choose the evaluation criteria, and the system automatically calculates the total score according to the selected criterion.

2. **Supplier Scoring**  
   In addition to drivers, the system also enables scoring for suppliers, using the defined criteria for their evaluation. The scoring process is similar to that of drivers, ensuring a consistent interface.

3. **Excel Spreadsheet Integration**  
   The system reads driver, supplier, and criteria data directly from an Excel spreadsheet. Information is automatically loaded, and any entered scores are saved in the same sheet, ensuring easy access to data.

4. **User-Friendly Graphical Interface**  
   Developed with PyQt5, the system offers a simple, intuitive graphical interface with separate tabs for drivers and suppliers, making it easy to use and navigate between different modules.

5. **Dark Theme**  
   To improve user experience, the system applies a dark theme, making prolonged use more comfortable.

**TECHNOLOGIES USED**

1. **Python**: Main language used for system development.
2. **PyQt5**: Library used to create the application's graphical interface.
3. **Pandas**: Used for reading and managing data from the Excel spreadsheet.
4. **Openpyxl**: Used to add new data and save entries in the Excel spreadsheet.

**HOW IT WORKS**

1. The user launches the system and selects the Excel file containing drivers, suppliers, and criteria data.
2. Through the interface, the user selects the relevant tab (Drivers or Suppliers).
3. The user selects the driver or supplier's name, chooses the evaluation criteria, and the score is automatically filled.
4. The system allows the user to save data directly in the Excel spreadsheet, updating the respective tab with new entries.
5. A confirmation message informs the user that the data has been successfully posted.

**CONCLUSION**

This system provides an efficient solution for posting and managing scores for drivers and suppliers. With a user-friendly interface and direct Excel file handling, the evaluation process becomes more practical and organized.
