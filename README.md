Student Invoice Generator
Overview
The Student Invoice Generator is a Java-based application that generates personalized PDF invoices for students based on data from an Excel file (Students.xlsx). Each invoice includes a dynamic QR code, student-specific details, a randomly generated billing amount, and tax calculations. The project leverages Apache POI for Excel processing, iText for PDF generation, and ZXing for QR code creation.

This tool is designed for scenarios where invoices need to be generated in bulk, such as for educational institutions or service providers tracking student-related charges.

Features
Dynamic Invoice Creation: Generates unique PDF invoices for each student listed in the Excel file.
QR Code Integration: Adds a scannable QR code to each invoice with student ID and amount details.
Excel Data Input: Reads student data (Name, ID, Phone) from Students.xlsx.
Tax Calculation: Applies an 18% IGST to a randomly generated base amount (₹500–₹2000).
Customizable Output: Saves invoices and QR codes in a results/ directory.
Error Handling: Skips invalid or empty rows in the Excel file and logs issues to the console.
Prerequisites
To run this project, ensure you have the following installed:

Java Development Kit (JDK): Version 8 or higher.
Maven: For dependency management (optional if using an IDE like IntelliJ with built-in Maven support).
IntelliJ IDEA (recommended) or any Java IDE.
A spreadsheet editor (e.g., Microsoft Excel, LibreOffice Calc) to edit Students.xlsx.
Dependencies
The project uses the following libraries (managed via Maven):

Apache POI: For reading and processing Excel files (poi and poi-ooxml).
iText: For generating PDF invoices (itextpdf).
ZXing: For creating QR codes (zxing).
