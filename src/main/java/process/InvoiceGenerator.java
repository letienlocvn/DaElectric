package process;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class InvoiceGenerator {

    public static void mapDataToTemplate(String templatePath,
                                         String outputPath,
                                         List<Map<String, String>> data) throws IOException {
        // Load the template
        try (FileInputStream fis = new FileInputStream((templatePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);

            // Starting row for duplication (adjust based on template structure)
            int startRow = 0; // Row where the template begins
            int endRow = 14; // Last row of the template section (both left and right parts)

            updatePlaceholders(sheet, startRow, endRow, data.get(0));

            // Step 3: Process additional rows
            // Track the current row for inserting data
            int currentRow = endRow + 1; // Leave room for the first template
            for (int i = 1; i < data.size(); i++) {
                copyAndFillTemplate(sheet, startRow, endRow, currentRow, data.get(i));
                currentRow += (endRow - startRow + 1); // Move the row pointer
            }

            // Save the updated file
            // Step 4: Save the workbook to the output file
            try (FileOutputStream fos = new FileOutputStream(new File(outputPath))) {
                workbook.write(fos);
            }
        }
    }

    // Function to update placeholders in the template section
    private static void updatePlaceholders(
            Sheet sheet,
            int startRow,
            int endRow,
            Map<String, String> rowData
    ) {
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    // Replace placeholders with corresponding values
                    for (Map.Entry<String, String> entry : rowData.entrySet()) {
                        if (cellValue.contains("{{" + entry.getKey() + "}}")) {
                            cell.setCellValue(cellValue.replace("{{" + entry.getKey() + "}}", entry.getValue()));
                        }
                    }
                }
            }
        }
    }

    // Function to copy template and fill data for subsequent rows
    private static void copyAndFillTemplate(
            Sheet sheet,
            int startRow,
            int endRow,
            int targetRow,
            Map<String, String> rowData
    ) {
        for (int i = 0; i <= (endRow - startRow); i++) {
            Row sourceRow = sheet.getRow(startRow + i);
            Row targetRowObj = sheet.createRow(targetRow + i);

            if (sourceRow != null) {
                copyRow(sourceRow, targetRowObj);
                updatePlaceholdersInRow(targetRowObj, rowData);
            }
        }
    }

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
        // Example data (replace with actual input)
        List<Map<String, String>> data = List.of(
                Map.of("index", "1", "fullName", "John Doe", "oldIndex", "100", "newIndex", "200", "unitPrice", "5000"),
                Map.of("index", "2", "fullName", "Jane Smith", "oldIndex", "150", "newIndex", "250", "unitPrice", "4500"),
                Map.of("index", "3",
                        "fullName", "Loc", "oldIndex", "150", "newIndex", "250", "unitPrice", "4500")
        );

        String templatePath = "data/final/input/HoaDon2023_Template.xlsx";
        String outputPath = "data/test/GeneratedInvoices.xlsx";

        mapDataToTemplate(templatePath, outputPath, data);

        System.out.println("Invoices generated successfully!");
    }
}

