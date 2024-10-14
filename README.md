# **Company Invoice Automation and PDF Export in Excel**

## **Project Overview**
This project automates the creation of company invoices using **Excel VBA (Visual Basic for Applications)**. It enables users to generate invoices automatically based on input data, format the invoice with a professional template, and export the final invoice as a **PDF**. This tool reduces manual work, enhances accuracy, and provides an easy way to export and share invoices.

## **Key Features**
- **Automated Invoice Generation:** Automatically populates invoice fields (customer name, products/services, prices, etc.) based on input data.
- **Customizable Invoice Template:** Easily modify the template to match your company’s branding.
- **Invoice Numbering:** Automatically generates and increments unique invoice numbers for each new invoice.
- **PDF Export:** Exports the generated invoice to a **PDF** file with the click of a button.
- **Error Handling:** Validates data before creating an invoice to ensure accuracy.

## **Technologies Used**
- **Microsoft Excel:** The primary tool for designing the invoice template and automating tasks.
- **VBA (Visual Basic for Applications):** Used to write macros that handle automation and PDF export.
- **PDF Export Capabilities:** Utilized through VBA to convert Excel sheets into PDF format.

## **Setup Instructions**
To set up and use this project on your local machine, follow these steps:

### **1. Enable Developer Mode in Excel**
- Open Excel.
- Go to **File** → **Options** → **Customize Ribbon**.
- In the right-hand pane, check the box labeled **Developer**.
- Click **OK** to display the Developer tab on your Ribbon.

### **2. Access the VBA Editor**
- Open the Excel workbook.
- Click the **Developer** tab on the Ribbon.
- Select **Visual Basic** to open the VBA editor.

### **3. Import or Create Macros**
- If a VBA module is provided, import it by going to **File** → **Import File** in the VBA editor and selecting the module.
- Otherwise, you can manually copy-paste the macro code into the **Modules** section in the VBA editor.

### **4. Save Workbook as Macro-Enabled**
- To ensure that the macros are saved correctly, go to **File** → **Save As**.
- Select **Excel Macro-Enabled Workbook (*.xlsm)** as the file type.

## **How to Use the Automation**
### **1. Input Invoice Data**
- Open the Excel workbook.
- Go to the sheet where invoice input fields are provided (e.g., Customer Details, Product/Service Details, Prices).
- Fill in the necessary information for generating an invoice (customer name, product details, unit price, quantity, etc.).

### **2. Generate Invoice**
- After filling out the data, click the **Generate Invoice** button (created using a macro).
- The macro will populate a formatted invoice template with the provided details and assign a unique invoice number.

### **3. Export Invoice as PDF**
- Once the invoice is generated, click the **Export as PDF** button.
- The macro will export the active invoice to a **PDF file** and save it to a specified location on your computer (which can be customized in the macro code).

### **4. Review and Share**
- The PDF invoice will be saved with a filename based on the invoice number (e.g., `Invoice_001.pdf`).
- You can then review the file or send it directly to the client.

## **Example Macros**
Here are some examples of the macros used in this project:

```vba
' Macro to Generate Invoice
Sub PDF()

Dim invoive_number As Long
Dim name As String
Dim file_path As String
Dim file_name As String

invoice_number = Range("G4")
name = Range("D5")
file_path = "use your path where you want to save the pdf file"

file_name = invoice_number & "_" & name

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, ignoreprintareas:=False, FileName:=file_path & file_name


End Sub

```

## **Folder Structure**
```
InvoiceAutomationProject/
│
├── InvoiceAutomation.xlsm     # The main Excel file with macros enabled
├── README.md                  # This readme file
└── VBA_Code_Module.bas        # Optional: Exported VBA code (if needed)
```

## **Requirements**
- **Microsoft Excel 2016** or newer.
- Basic knowledge of Excel and VBA (Visual Basic for Applications).
- PDF viewer (to view the exported PDF files).

## **Troubleshooting**
1. **Macros not working:**
   - Ensure macros are enabled. When opening the file, click **Enable Macros** if prompted.
2. **PDF not exporting:**
   - Check that the file path in the macro is correct and that you have permission to save files in that location.
3. **Invoice number not incrementing:**
   - Review the logic in the macro responsible for generating invoice numbers. Ensure the cell reference for storing the last invoice number is correct.

## **Contributing**
If you'd like to contribute to this project, feel free to submit pull requests or report issues. Contributions are welcome to improve features, add customization, or enhance the user experience.

---

Let me know if you'd like to add any specific details or instructions!

![Untitled design](https://github.com/user-attachments/assets/005b571d-1a68-46d0-840d-6b5ef6886197)
