import model.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelReaderTest {
    @Test
    void testReadInputFile_withValidData() throws IOException {
        // Create a sample Excel file
        String testFilePath = "data/test/test_ElectricManagement.xlsx";
        createSampleExcelFile(testFilePath);

        // Call the function
        List<ExcelReader.ElectricBillRecord> records = ExcelReader.readInputFile(testFilePath);

        // Validate results
        assertEquals(2, records.size(), "Should have 2 records");

        ExcelReader.ElectricBillRecord record1 = records.get(0);
        assertEquals(1, record1.index);
        assertEquals("John Doe", record1.customerName);
        assertEquals(100.0, record1.oldIndex);
        assertEquals(150.0, record1.newIndex);
        assertEquals(185.0, record1.totalPayment);

        ExcelReader.ElectricBillRecord record2 = records.get(1);
        assertEquals(2, record2.index);
        assertEquals("Jane Smith", record2.customerName);
        assertEquals(200.0, record2.oldIndex);
        assertEquals(300.0, record2.newIndex);
        assertEquals(360.0, record2.totalPayment);

        // Cleanup
        new File(testFilePath).delete();
    }

    @Test
    void testReadInputFile_withEmptyFile() throws IOException {
        // Create an empty Excel file
        String testFilePath = "empty_ElectricManagement.xlsx";
        createEmptyExcelFile(testFilePath);

        // Call the function
        List<ExcelReader.ElectricBillRecord> records = ExcelReader.readInputFile(testFilePath);

        // Validate results
        assertTrue(records.isEmpty(), "Should return an empty list for an empty file");

        // Cleanup
        new File(testFilePath).delete();
    }

    private void createSampleExcelFile(String filePath) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(filePath)) {
            XSSFSheet sheet = workbook.createSheet();

            // Header Row
            XSSFRow header = sheet.createRow(0);
            header.createCell(0).setCellValue("STT");
            header.createCell(1).setCellValue("Họ Tên");
            header.createCell(2).setCellValue("Chỉ số cũ");
            header.createCell(3).setCellValue("Chỉ số mới");
            header.createCell(4).setCellValue("Số trong tháng");
            header.createCell(5).setCellValue("Đơn giá");
            header.createCell(6).setCellValue("Thành tiền");
            header.createCell(7).setCellValue("Công ghi điện");
            header.createCell(8).setCellValue("Tổng thanh toán");

            // Data Rows
            XSSFRow row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(1);
            row1.createCell(1).setCellValue("John Doe");
            row1.createCell(2).setCellValue(100);
            row1.createCell(3).setCellValue(150);
            row1.createCell(4).setCellValue(50);
            row1.createCell(5).setCellValue(3.5);
            row1.createCell(6).setCellValue(175);
            row1.createCell(7).setCellValue(10);
            row1.createCell(8).setCellValue(185);

            XSSFRow row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue(2);
            row2.createCell(1).setCellValue("Jane Smith");
            row2.createCell(2).setCellValue(200);
            row2.createCell(3).setCellValue(300);
            row2.createCell(4).setCellValue(100);
            row2.createCell(5).setCellValue(3.5);
            row2.createCell(6).setCellValue(350);
            row2.createCell(7).setCellValue(10);
            row2.createCell(8).setCellValue(360);

            workbook.write(fos);
        }
    }

    private void createEmptyExcelFile(String filePath) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.createSheet();
            workbook.write(fos);
        }
    }
}
