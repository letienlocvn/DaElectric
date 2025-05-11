package process;

import model.ExcelReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * With version 1.2:
 * Complete the generated invoices
 * Changes: Code clean
 */
public class InvoiceGenerator {

    public static void mapDataToTemplate(String templatePath,
                                         String outputPath,
                                         List<Map<String, Object>> data) throws IOException {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int startRow = 0;
            int endRow = 11;
            int currentRow = endRow + 1;

            // Process each data row and generate the invoice
            for (Map<String, Object> rowData : data) {
                generateInvoice(sheet, startRow, endRow, rowData, currentRow);
                currentRow += (endRow - startRow + 1);
            }

            saveWorkbook(workbook, outputPath);
        }
    }

    private static void generateInvoice(Sheet sheet,
                                        int startRow,
                                        int endRow,
                                        Map<String, Object> rowData,
                                        int currentRow) {
        int generatedRowStart = currentRow;
        int generatedRowEnd = currentRow + (endRow - startRow);

        // Copy rows from template
        copyTemplateRows(sheet, startRow, endRow, currentRow);

        // Update placeholders in the generated rows
        updatePlaceholders(sheet, generatedRowStart, generatedRowEnd, rowData);
    }

    private static void copyTemplateRows(Sheet sheet, int startRow, int endRow, int currentRow) {
        for (int i = startRow; i <= endRow; i++) {
            Row sourceRow = sheet.getRow(i);
            Row targetRow = sheet.createRow(currentRow++);
            if (sourceRow != null) {
                copyRow(sheet, sourceRow, targetRow);
            }
        }
    }

    private static void updatePlaceholders(Sheet sheet, int startRow, int endRow, Map<String, Object> rowData) {
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    replacePlaceholders(cell, cellValue, rowData);
                }
            }
        }
    }

    private static void replacePlaceholders(Cell cell, String cellValue, Map<String, Object> rowData) {
        for (Map.Entry<String, Object> entry : rowData.entrySet()) {
            String placeholder = "{{" + entry.getKey() + "}}";
            if (cellValue.contains(placeholder)) {
                if (entry.getValue() instanceof String) {
                    cell.setCellValue(cellValue.replace(placeholder, (String) entry.getValue()));
                } else if (entry.getValue() instanceof Number) {
                    cell.setCellValue(((Number) entry.getValue()).doubleValue());
                }
            }
        }
    }

    private static void copyRow(Sheet sheet, Row sourceRow, Row targetRow) {
        if (sourceRow == null || targetRow == null) return;

        targetRow.setHeight(sourceRow.getHeight());
        for (Cell sourceCell : sourceRow) {
            Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());
            copyCell(sourceCell, targetCell);
        }

        copyMergedRegions(sheet, sourceRow, targetRow);
    }

    private static void copyMergedRegions(Sheet sheet, Row sourceRow, Row targetRow) {
        int targetRowNum = targetRow.getRowNum();
        int sourceRowNum = sourceRow.getRowNum();

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() == sourceRowNum && region.getLastRow() == sourceRowNum) {
                CellRangeAddress newRegion = new CellRangeAddress(
                        targetRowNum, targetRowNum, region.getFirstColumn(), region.getLastColumn());
                sheet.addMergedRegion(newRegion);
            }
        }
    }

    private static void copyCell(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) return;

        targetCell.setCellStyle(sourceCell.getCellStyle());

        switch (sourceCell.getCellType()) {
            case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
            case NUMERIC -> targetCell.setCellValue(sourceCell.getNumericCellValue());
            case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
            case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
            case BLANK -> targetCell.setBlank();
        }
    }

    private static void saveWorkbook(XSSFWorkbook workbook, String outputPath) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        }
    }

    private static List<Map<String, Object>> convertToTemplateData(List<ExcelReader.ElectricBillRecord> records) {
        List<Map<String, Object>> data = new ArrayList<>();
        for (ExcelReader.ElectricBillRecord record : records) {
            Map<String, Object> rowData = new HashMap<>();
            rowData.put("index", record.index);
            rowData.put("fullName", record.customerName);
            rowData.put("oldIndex", record.oldIndex);
            rowData.put("newIndex", record.newIndex);
            rowData.put("unitsInMonth", record.unitsInMonth);
            rowData.put("unitPrice", record.unitPrice);
            rowData.put("totalPayment", record.totalPayment);
            data.add(rowData);
        }
        return data;
    }

    public static void main(String[] args) throws IOException {
        String inputFilePath = "data/final/input/ElectricityManagement.xlsx";
        String templatePath = "data/final/input/HoaDon2023_Template.xlsx";
        String outputPath = "data/final/output/GeneratedInvoices.xlsx";

        // Read input data
        List<ExcelReader.ElectricBillRecord> records = ExcelReader.readInputFile(inputFilePath);

        // Prepare data for template
        List<Map<String, Object>> data = convertToTemplateData(records);

        // Generate invoices
        mapDataToTemplate(templatePath, outputPath, data);

        System.out.println("Invoices generated successfully!");
    }
}

