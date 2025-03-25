package org.TMSS;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Generates personalized PDF invoices for students based on Excel data.
 * Invoices include dynamic QR codes, student-specific details, and configurable billing amounts.
 */
public class StudentInvoiceGenerator {

    private static final String RESULTS_DIR = "results";
    private static final String EXCEL_FILE_PATH = "Students.xlsx";
    private static final String LOGO_PATH = "src/main/resources/TMSS LOGO.jpg";
    private static final String INVOICE_DATE = LocalDate.now().format(DateTimeFormatter.ofPattern("dd-MMM-yyyy"));
    private static final double IGST_RATE = 0.18;

    public static void main(String[] args) {
        ensureDirectoryExists(RESULTS_DIR);

        List<String[]> students = loadStudentData(EXCEL_FILE_PATH);
        if (students.isEmpty()) {
            System.out.println("No valid student data found in " + EXCEL_FILE_PATH);
            return;
        }

        for (String[] student : students) {
            String fileName = createInvoiceFileName(student[1]);
            generateInvoice(fileName, student);
        }
    }

    // Ensures the output directory exists
    private static void ensureDirectoryExists(String directory) {
        File dir = new File(directory);
        if (!dir.exists()) {
            dir.mkdir();
        }
    }

    // Loads student data (name, ID, phone) from Excel, skipping invalid rows
    private static List<String[]> loadStudentData(String filePath) {
        List<String[]> students = new ArrayList<>();
        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);
            boolean skipHeader = true;

            for (Row row : sheet) {
                if (skipHeader) {
                    skipHeader = false;
                    continue;
                }

                // Skip rows that are physically empty
                if (isRowEmpty(row)) {
                    continue;
                }

                String name = getCellValue(row, 0);
                String studentId = getCellValue(row, 1);
                String phone = getCellValue(row, 2);

                // Debugging output to see what’s being read
                System.out.println("Row data - Name: '" + name + "', ID: '" + studentId + "', Phone: '" + phone + "'");

                // Only add students with non-empty name and studentId
                if (isValidStudent(name, studentId)) {
                    students.add(new String[]{name, studentId, phone});
                } else {
                    System.out.println("Skipping invalid row - Name: '" + name + "', ID: '" + studentId + "'");
                }
            }
        } catch (Exception e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
        }
        return students;
    }

    // Checks if a row is completely empty
    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (cell != null && !cell.toString().trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    // Validates that a student has required fields
    private static boolean isValidStudent(String name, String studentId) {
        return name != null && !name.trim().isEmpty() &&
                studentId != null && !studentId.trim().isEmpty();
    }

    // Safely retrieves cell value as a trimmed string
    private static String getCellValue(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        return cell.toString().trim();
    }

    // Creates a unique invoice filename based on student ID
    private static String createInvoiceFileName(String studentId) {
        String safeId = studentId.replace(" ", "_");
        return RESULTS_DIR + "/Invoice_" + safeId + ".pdf";
    }

    // Generates a PDF invoice for a student
    private static void generateInvoice(String filePath, String[] student) {
        Document document = new Document();
        try {
            PdfWriter.getInstance(document, new FileOutputStream(filePath));
            document.open();

            int baseAmount = new Random().nextInt(1501) + 500; // Generate once per invoice

            addCompanyLogo(document);
            addInvoiceHeader(document, student[2]);
            addCompanyDetails(document);
            addBillingInfo(document, student[0]);
            addInvoiceItems(document, baseAmount);
            addPaymentSummary(document, baseAmount);
            addQRCode(document, student[1], baseAmount); // Use studentId for QR
            addBankDetails(document);
            addSignature(document);

            document.close();
            System.out.println("Generated invoice: " + filePath + " with amount Rs. " + baseAmount);
        } catch (Exception e) {
            System.err.println("Error generating invoice: " + e.getMessage());
        }
    }

    // Adds the company logo
    private static void addCompanyLogo(Document document) throws Exception {
        Image logo = Image.getInstance(LOGO_PATH);
        logo.scaleToFit(200, 50);
        logo.setAlignment(Element.ALIGN_CENTER);
        document.add(logo);
    }

    // Adds invoice header with dynamic date and phone-based invoice number
    private static void addInvoiceHeader(Document document, String phone) throws DocumentException {
        Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16);
        Font normalFont = FontFactory.getFont(FontFactory.HELVETICA, 12);

        document.add(new Paragraph("INVOICE", titleFont));
        document.add(Chunk.NEWLINE);
        document.add(new Paragraph("Invoice No.: TMSS/2024-2025/DPSK/INV/" + phone, normalFont));
        document.add(new Paragraph("Invoice Date: " + INVOICE_DATE, normalFont));
        document.add(Chunk.NEWLINE);
    }

    // Adds static company details
    private static void addCompanyDetails(Document document) throws DocumentException {
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
        Font normalFont = FontFactory.getFont(FontFactory.HELVETICA, 12);

        document.add(new Paragraph("From:", boldFont));
        document.add(new Paragraph("TechnoMedia Software Solutions Pvt. Ltd.", normalFont));
        document.add(new Paragraph("CV Ramannagar, Bangalore, Karnataka - 560075", normalFont));
        document.add(new Paragraph("Email: info@technomediasoft.com", normalFont));
        document.add(Chunk.NEWLINE);
    }

    // Adds dynamic billing info based on student name
    private static void addBillingInfo(Document document, String studentName) throws DocumentException {
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
        document.add(new Paragraph("Bill To: " + studentName, boldFont));
        document.add(Chunk.NEWLINE);
    }

    // Adds invoice items with a consistent amount
    private static void addInvoiceItems(Document document, int baseAmount) throws DocumentException {
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);

        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);
        table.addCell(new PdfPCell(new Phrase("Sl. No.", boldFont)));
        table.addCell(new PdfPCell(new Phrase("Description", boldFont)));
        table.addCell(new PdfPCell(new Phrase("Amount (Rs.)", boldFont)));
        table.addCell("1");
        table.addCell("RouteAlert charges for Nov-2024\nNumber Of Students: 748");
        table.addCell(String.valueOf(baseAmount));
        document.add(table);
        document.add(Chunk.NEWLINE);
    }

    // Adds payment summary with consistent amount and tax calculation
    private static void addPaymentSummary(Document document, int baseAmount) throws DocumentException {
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
        Font normalFont = FontFactory.getFont(FontFactory.HELVETICA, 12);

        double discount = 0.00;
        double subtotal = baseAmount - discount;
        double igst = subtotal * IGST_RATE;
        double totalAmount = subtotal + igst;

        document.add(new Paragraph("Other Charges:", boldFont));
        document.add(new Paragraph("Discount: Rs. " + String.format("%.2f", discount), normalFont));
        document.add(new Paragraph("Subtotal: Rs. " + String.format("%.2f", subtotal), normalFont));
        document.add(new Paragraph("IGST @ 18%: Rs. " + String.format("%.2f", igst), normalFont));
        document.add(new Paragraph("Total: Rs. " + String.format("%.2f", totalAmount) + " only", boldFont));
        document.add(Chunk.NEWLINE);
    }

    // Adds a dynamic QR code based on student ID and amount
    private static void addQRCode(Document document, String studentId, int baseAmount) throws Exception {
        String qrFileName = RESULTS_DIR + "/QR_" + studentId + ".png";
        createQRCode(studentId, baseAmount, qrFileName);
        Image qrImage = Image.getInstance(qrFileName);
        qrImage.scaleToFit(100, 100);
        qrImage.setAlignment(Element.ALIGN_CENTER);
        document.add(qrImage);
        document.add(Chunk.NEWLINE);
    }

    // Creates a QR code with student ID and amount details
    private static void createQRCode(String studentId, int baseAmount, String fileName) {
        int width = 200;
        int height = 200;
        try {
            String qrContent = "Invoice for Student ID: " + studentId + "\nAmount: Rs. " + baseAmount;
            BitMatrix bitMatrix = new MultiFormatWriter().encode(
                    qrContent, BarcodeFormat.QR_CODE, width, height
            );
            MatrixToImageWriter.writeToPath(bitMatrix, "PNG", Paths.get(fileName));
            System.out.println("QR Code created: " + fileName);
        } catch (Exception e) {
            System.err.println("Error generating QR code: " + e.getMessage());
        }
    }

    // Adds static bank details
    private static void addBankDetails(Document document) throws DocumentException {
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
        Font normalFont = FontFactory.getFont(FontFactory.HELVETICA, 12);

        document.add(new Paragraph("Bank Details:", boldFont));
        document.add(new Paragraph("TechnoMedia Software Solutions Pvt. Ltd.", normalFont));
        document.add(new Paragraph("Bank: Indian Overseas Bank", normalFont));
        document.add(new Paragraph("Branch: Bangalore-560075", normalFont));
        document.add(Chunk.NEWLINE);
    }

    // Adds the signature section
    private static void addSignature(Document document) throws DocumentException {
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
        Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16);

        document.add(new Paragraph("For TechnoMedia Software Solutions Pvt. Ltd.", boldFont));
        document.add(new Paragraph("Authorized Signatory", titleFont));
    }
}