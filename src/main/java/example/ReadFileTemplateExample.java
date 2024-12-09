package example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadFileTemplateExample {

    public static void main(String[] args) {
        String filePath = "D:\\Loclt\\Project_out\\loclt7\\DaElectric\\src\\HoaDon2023.xlsx";

        try {
            readAllData(filePath);
        } catch (IOException e) {
            System.err.println("Error reading the Excel file: " + e.getMessage());
        }
    }

    public static void readAllData(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Lặp qua từng sheet
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                System.out.println("Sheet Name: " + sheet.getSheetName());

                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (row.getRowNum() == 0) continue;
                        String name = row.getCell(1).getStringCellValue(); // Column B: "Họ Tên"
                        double oldIndex = row.getCell(2).getNumericCellValue(); // Column C: "Chỉ số cũ"
                        double newIndex = row.getCell(3).getNumericCellValue(); // Column D: "Chỉ số mới"
                        double unitPrice = row.getCell(5).getNumericCellValue(); // Column F: "Đơn giá"
                        double totalPayment = row.getCell(8).getNumericCellValue(); // Column I: "Tổng thanh toán"

                        // Clone template and replace placeholders
                        replacePlaceholder(sheet, "{{id}}", name);
                        replacePlaceholder(sheet, "{{fullName}}", String.valueOf(oldIndex));
                        replacePlaceholder(sheet, "{{oldIndex}}", String.valueOf(newIndex));
                        replacePlaceholder(sheet, "{{newIndex}}", String.valueOf(unitPrice));
                        replacePlaceholder(sheet, "{{unitPrice}}", String.valueOf(totalPayment));
                    }
                    System.out.println(); // Xuống dòng sau mỗi hàng
                }
            }
        }
    }
    private static void replacePlaceholder(Sheet sheet, String placeholder, String value) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().contains(placeholder)) {
                    cell.setCellValue(cell.getStringCellValue().replace(placeholder, value));
                }
            }
        }
    }


}
