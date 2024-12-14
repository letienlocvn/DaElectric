**Project Overview**
This Java application automates the generation of electricity invoices using Excel files. It reads billing data from an input file, processes it, and populates a pre-designed Excel template to generate invoices for each customer. The goal is to eliminate manual data entry while preserving the template's formatting and structure.

**Input File**
* **Structure:** Single sheet with predefined columns:
    * STT (Index) - Integer
    * Họ Tên (Customer Name) - String
    * Chỉ số cũ (Old Index) - Double
    * Chỉ số mới (New Index) - Double
    * Số trong tháng (Units in Month) - Double
    * Đơn giá (Unit Price) - Double
    * Thành tiền (Amount) - Double
    * Công ghi điện (Recording Fee) - Double
    * Tổng thanh toán (Total Payment) - Double
    * Volume: Roughly 100 rows per file
* **Purpose:** Each row represents one invoice to be generated.

**Invoice Template (HoaDon2023_Template.xlsx)**

* **Structure:** Detailed Excel template with pre-designed invoice format.
* **Includes:** Dynamic placeholders like `{{index}}`, `{{fullName}}`, `{{oldIndex}}`, etc.
* **Preserves:** Complex formatting, merged cells, alignment, and embedded formulas.
* **Sections:** Split into two labeled sections ("Liên") with repeated data for different formats.

**Output Requirements**
* **Output File:** A new Excel file (GeneratedInvoices.xlsx) containing invoices for all customers.
* **Output Format:**
    * Each customer's invoice uses a duplicated template section.
    * Dynamic placeholders replaced with corresponding values from the input file.
    * Template styles, formatting, and formulas are preserved.
    * All invoices appear sequentially in one sheet.

**Key Features**

* **Input Parsing:** Reads data from the input Excel file using Apache POI and creates a list of `ElectricBillRecord` objects for each row.
* **Invoice Generation:**
    * Duplicates the template for each customer.
    * Replaces placeholders with actual values, preserving data types (numeric/string).
    * Maintains template formatting, styles, and formulas.
* **Output File:** Saves generated invoices to a single consolidated Excel file.
* **Error Handling:**
    * Handles missing/mismatched data logically.
    * Implements error logging for skipped rows or processing issues.

**Technology Stack**
* Java Libraries: Apache POI for Excel file handling
* Tools: XSSFWorkbook and XSSFSheet for reading/writing XLSX files
* Dynamic placeholder replacement logic
**Current Progress**

* **Input Parsing:** Successful reading of data into `ElectricBillRecord` list with data type preservation.
* **Template Processing:** Invoice generation with template duplication, placeholder replacement, and formatting preservation.
**Output File:** Generation of a consolidated Excel file with all invoices.

**Pending Items/Features**

* **Testing:** Validate output using real-world data and scenarios. Verify formatting, alignment, and formula behavior.
* **Optimization:** Ensure scalability for larger datasets. Refactor code for readability and maintainability.
* **Error Handling:** Handle missing data in the input file and improve logging for debugging and reporting.
* **Future Enhancements:**
    * PDF generation for printing.
    * Dynamic template sheet names or multiple templates.
    * User interface for non-technical users.

**Note:** This Readme.md provides a basic overview of the project.  For detailed documentation and code explanations, refer to the project's source code itself.
