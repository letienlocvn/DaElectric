package process;

import model.ExcelReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * With version 1.1:
 * Change the business - Generated template first -> Update Placeholder later
 */
public class InvoiceGenerator {

    public static void mapDataToTemplate(String templatePath,
                                         String outputPath,
                                         List<Map<String, Object>> data) throws IOException {
        // Load the template
        FileInputStream fis = new FileInputStream(new File(templatePath));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);

        // Define the template row range
        int startRow = 0;  // Adjust based on template
        int endRow = 13;   // Last row of the template section

        int currentRow = endRow + 1; // Start writing after the template section

        for (Map<String, Object> rowData : data) {
            // Generate template for each record
            int generatedRowStart = currentRow;
            int generatedRowEnd = currentRow + (endRow - startRow);

            for (int i = startRow; i <= endRow; i++) {
                Row sourceRow = sheet.getRow(i);
                Row targetRow = sheet.createRow(currentRow++);

                if (sourceRow != null) {
                    copyRow(sourceRow, targetRow);
                }
            }

            // Update placeholders in the generated rows
            updatePlaceholders(sheet, generatedRowStart, generatedRowEnd, rowData);
        }

        // Save the updated file
        FileOutputStream fos = new FileOutputStream(new File(outputPath));
        workbook.write(fos);
        fos.close();
        workbook.close();
        fis.close();
    }

    // Function to update placeholders in the template section
    private static void updatePlaceholders(
            Sheet sheet,
            int startRow,
            int endRow,
            Map<String, Object> rowData
    ) {
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    // Replace placeholders with corresponding values
                    for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                        if (cellValue.contains("{{" + entry.getKey() + "}}")) {
                            if (entry.getValue() instanceof String) {
                                cell.setCellValue(cellValue.replace("{{" + entry.getKey() + "}}", (String) entry.getValue()));
                            } else if (entry.getValue() instanceof Number) {
                                cell.setCellValue(((Number) entry.getValue()).doubleValue());
                            }
                        }
                    }
                }
            }
        }
    }

    // Function to copy template and fill data for subsequent rows

    // Function to copy styles and content of a row
    private static void copyRow(Row sourceRow, Row targetRow) {
        for (Cell sourceCell : sourceRow) {
            Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());
            copyCell(sourceCell, targetCell);
        }
    }

    // Function to copy individual cell content and style
    private static void copyCell(Cell sourceCell, Cell targetCell) {
        targetCell.setCellStyle(sourceCell.getCellStyle());

        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }

    // Function to update placeholders in a single row
    private static void updatePlaceholdersInRow(Row row, Map<String, String> rowData) {
        for (Cell cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                String cellValue = cell.getStringCellValue();
                for (Map.Entry<String, String> entry : rowData.entrySet()) {
                    if (cellValue.contains("{{" + entry.getKey() + "}}")) {
                        cell.setCellValue(cellValue.replace("{{" + entry.getKey() + "}}", entry.getValue()));
                    }
                }
            }
        }
    }

    public static void main(String[] args) throws IOException {
        // Define file paths
        String inputFilePath = "data/final/input/ElectricityManagement.xlsx";
        String templatePath = "data/final/input/HoaDon2023_Template.xlsx";
        String outputPath = "data/final/output/GeneratedInvoices.xlsx"; // Replace with your desired output file path

        // Read input data from Excel file
        List<ExcelReader.ElectricBillRecord> records = ExcelReader.readInputFile(inputFilePath);
        // Convert records to a format compatible with the template
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

        mapDataToTemplate(templatePath, outputPath, data);

        System.out.println("Invoices generated successfully!");
    }
}

