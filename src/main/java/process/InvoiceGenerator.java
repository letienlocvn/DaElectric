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
                                         List<Map<String, Object>> data,
                                         int startRow,
                                         int endRow
    ) throws IOException {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int templateHeight = endRow - startRow + 1;
            int currentRow = endRow + 1;

            // Thu thập danh sách các vùng gộp ô trong Template gốc một lần duy nhất
            List<CellRangeAddress> templateMergedRegions = getTemplateMergedRegions(sheet, startRow, endRow);

            // Process each data row and generate the invoice
            for (Map<String, Object> rowData : data) {
                generateInvoice(sheet, startRow, endRow, rowData, currentRow, templateMergedRegions);
                currentRow += templateHeight;
            }

            saveWorkbook(workbook, outputPath);
        }
    }

    private static void generateInvoice(Sheet sheet,
                                        int startRow,
                                        int endRow,
                                        Map<String, Object> rowData,
                                        int currentRow,
                                        List<CellRangeAddress> templateMergedRegions) {
        int generatedRowStart = currentRow;
        int generatedRowEnd = currentRow + (endRow - startRow);

        // 1. Copy dữ liệu dòng và style
        copyTemplateRows(sheet, startRow, endRow, currentRow);

        // 2. Copy các vùng gộp ô (Merged Regions) theo khối
        copyBlockMergedRegions(sheet, startRow, generatedRowStart, templateMergedRegions);

        // 3. Thay thế Placeholder
        updatePlaceholders(sheet, generatedRowStart, generatedRowEnd, rowData);
    }

    /**
     * Lấy danh sách các vùng gộp ô nằm trọn trong khối Template gốc
     */
    private static List<CellRangeAddress> getTemplateMergedRegions(Sheet sheet, int startRow, int endRow) {
        List<CellRangeAddress> regions = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() >= startRow && region.getLastRow() <= endRow) {
                regions.add(region);
            }
        }
        return regions;
    }

    /**
     * Sao chép các vùng gộp ô từ Template sang vị trí mới
     */
    private static void copyBlockMergedRegions(Sheet sheet,
                                               int startRow,
                                               int targetStartRow,
                                               List<CellRangeAddress> templateRegions) {
        int rowOffset = targetStartRow - startRow;

        for (CellRangeAddress region : templateRegions) {
            CellRangeAddress newRegion = new CellRangeAddress(
                region.getFirstRow() + rowOffset,
                region.getLastRow() + rowOffset,
                region.getFirstColumn(),
                region.getLastColumn()
            );

            // Bỏ qua nếu vùng này đã bị gộp ô trước đó (tránh văng lỗi IllegalStateException)
            if (!isRegionMerged(sheet, newRegion)) {
                sheet.addMergedRegion(newRegion);
            }
        }
    }

    private static boolean isRegionMerged(Sheet sheet, CellRangeAddress targetRegion) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress existing = sheet.getMergedRegion(i);
            if (existing.intersects(targetRegion)) {
                return true;
            }
        }
        return false;
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
                Object value = entry.getValue();
                if (value == null) {
                    cellValue = cellValue.replace(placeholder, "");
                } else if (cellValue.equals(placeholder) && value instanceof Number) {
                    // Nếu ô chỉ chứa duy nhất 1 placeholder và giá trị là Số -> Giữ nguyên kiểu Numeric
                    cell.setCellValue(((Number) value).doubleValue());
                    return;
                } else {
                    // Trường hợp chuỗi kết hợp (Ví dụ: "Tổng tiền: {{totalPayment}} VNĐ")
                    cellValue = cellValue.replace(placeholder, String.valueOf(value));
                }
            }
        }
        cell.setCellValue(cellValue);
    }

    private static void copyTemplateRows(Sheet sheet, int startRow, int endRow, int currentRow) {
        for (int i = startRow; i <= endRow; i++) {
            Row sourceRow = sheet.getRow(i);
            Row targetRow = sheet.createRow(currentRow++);
            if (sourceRow != null) {
                copyRow(sourceRow, targetRow);
            }
        }
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        if (sourceRow == null || targetRow == null) return;

        targetRow.setHeight(sourceRow.getHeight());
        for (Cell sourceCell : sourceRow) {
            Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());
            copyCell(sourceCell, targetCell);
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
            rowData.put("category", record.category);
            rowData.put("oldIndex", record.oldIndex);
            rowData.put("newIndex", record.newIndex);
            rowData.put("unitsInMonth", record.unitsInMonth);
            rowData.put("unitPrice", record.unitPrice);
            rowData.put("tax", record.tax);
            rowData.put("totalPayment", record.totalPayment);
            data.add(rowData);
        }
        return data;
    }

    public static void main(String[] args) throws IOException {
        // Danh
//        String inputFilePath = "data/final/input/ElectricityManagement.xlsx";
//        String templatePath = "data/final/input/HoaDon2023_Template.xlsx";
        String outputPath = "data/final/output/GeneratedInvoices.xlsx";

        // Loi
        String inputFilePath = "data/final/input/TienDien_loi.xlsx";
        String templatePath = "data/final/input/HoaDonLoi_Template.xlsx";

        // Vị trí bắt đầu và kết thúc của Template hóa đơn (0-indexed)
        // Lưu ý: Đảm bảo startRow và endRow khớp đúng với số dòng của Template HoaDonLoi_Template.xlsx
        int startRow = 0;
        int endRow = 13;

        // Read input data
        List<ExcelReader.ElectricBillRecord> records = ExcelReader.readInputFile(inputFilePath);

        // Prepare data for template
        List<Map<String, Object>> data = convertToTemplateData(records);

        // Generate invoices
        mapDataToTemplate(templatePath, outputPath, data, startRow, endRow);

        System.out.println("Invoices generated successfully!");
    }
}

