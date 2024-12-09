package process;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelProcessor {
    public void processFiles(String managementFilePath, String templateFilePath, String outputFolderPath) throws IOException {
        // Load management file
        FileInputStream managementFile = new FileInputStream(managementFilePath);
        Workbook managementWorkbook = new XSSFWorkbook(managementFile);
        Sheet managementSheet = managementWorkbook.getSheetAt(0);

        // Load template file
        FileInputStream templateFile = new FileInputStream(templateFilePath);
        Workbook templateWorkbook = new XSSFWorkbook(templateFile);

        // Process each row in management file
        for (Row row : managementSheet) {
            if (row.getRowNum() == 0) continue; // Skip header row

            String name = row.getCell(1).getStringCellValue(); // Column B: "Họ Tên"
            double oldIndex = row.getCell(2).getNumericCellValue(); // Column C: "Chỉ số cũ"
            double newIndex = row.getCell(3).getNumericCellValue(); // Column D: "Chỉ số mới"
            double unitPrice = row.getCell(5).getNumericCellValue(); // Column F: "Đơn giá"
            double totalPayment = row.getCell(8).getNumericCellValue(); // Column I: "Tổng thanh toán"

            // Clone template and replace placeholders
            Sheet templateSheet = templateWorkbook.cloneSheet(0); // Clone template for each customer
            replacePlaceholder(templateSheet, "{{Tên}}", name);
            replacePlaceholder(templateSheet, "{{Chỉ số cũ}}", String.valueOf(oldIndex));
            replacePlaceholder(templateSheet, "{{Chỉ số mới}}", String.valueOf(newIndex));
            replacePlaceholder(templateSheet, "{{Đơn giá}}", String.valueOf(unitPrice));
            replacePlaceholder(templateSheet, "{{Tổng tiền thanh toán}}", String.valueOf(totalPayment));
        }

        // Save updated file
        FileOutputStream outputFile = new FileOutputStream(outputFolderPath);
        templateWorkbook.write(outputFile);

        // Close resources
        managementWorkbook.close();
        templateWorkbook.close();
        managementFile.close();
        outputFile.close();
    }

    private void replacePlaceholder(Sheet sheet, String placeholder, String value) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().contains(placeholder)) {
                    cell.setCellValue(cell.getStringCellValue().replace(placeholder, value));
                }
            }
        }
    }
}
