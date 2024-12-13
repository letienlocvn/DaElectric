package example;

import model.Customer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static example.ReadExcelExample.*;

public class CloneTemplate {
    public static void main(String[] args) {

        try {
            List<Customer> customers = readCustomersFromExcel(filePath);

            // Mở file template
            Workbook workbook = new XSSFWorkbook(new FileInputStream(invoiceTemplatePath));
            Sheet templateSheet = workbook.getSheetAt(0);
            int startRowIndex = templateSheet.getLastRowNum() + 1;
            Map<CellStyle, CellStyle> styleCache = new HashMap<>();
            for (Customer customer : customers) {
                // Update dữ liệu trong sheet template
                updateTemplateWithCustomerData(templateSheet, customer);
                appendInvoiceForCustomer(
                        templateSheet,
                        templateSheet,
                        customer,
                        startRowIndex,
                        workbook,
                        styleCache
                );
            }
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            // Lưu file mới
            FileOutputStream fileOut = new FileOutputStream(outputPath);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Invoices generated successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void appendInvoiceForCustomer(Sheet outputSheet, Sheet templateSheet, Customer customer, int startingRow, Workbook workbook, Map<CellStyle, CellStyle> styleCache) {
        // Lấy số lượng hàng trong template
        int templateRowCount = templateSheet.getLastRowNum() + 1;

        // Duyệt từng hàng trong template và copy vào sheet output
        for (int i = 0; i < templateRowCount; i++) {
            Row templateRow = templateSheet.getRow(i);
            Row newRow = outputSheet.createRow(startingRow + i);

            if (templateRow != null) {
                copyRow(templateRow, newRow, workbook, styleCache);
            }
        }

        // Xử lý dòng đầu tiên cho mỗi khách hàng
        Row firstRow = outputSheet.getRow(startingRow);

        if (firstRow != null) {
            mergeAndCenterCells(firstRow, workbook);
        }

        // Thay thế dữ liệu khách hàng trong các hàng mới
        for (int i = 0; i < templateRowCount; i++) {
            Row newRow = outputSheet.getRow(startingRow + i);

            if (newRow != null) {
                for (Cell cell : newRow) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        cell.setCellValue(replacePlaceholders(cellValue, customer));
                    }
                }
            }
        }
    }


    private static void mergeAndCenterCells(Row row, Workbook workbook) {
        Sheet sheet = row.getSheet();
        int rowIndex = row.getRowNum();

        // Gộp A14 và B14 nếu chưa được gộp
        CellRangeAddress range1 = new CellRangeAddress(rowIndex, rowIndex, 0, 1);
        if (!isMergedRegionAlreadyExists(sheet, range1)) {
            sheet.addMergedRegion(range1);
            setCellStyleAndAlignment(sheet, range1, workbook);
        }

        // Gộp D14 và E14 nếu chưa được gộp
        CellRangeAddress range2 = new CellRangeAddress(rowIndex, rowIndex, 3, 4);
        if (!isMergedRegionAlreadyExists(sheet, range2)) {
            sheet.addMergedRegion(range2);
            setCellStyleAndAlignment(sheet, range2, workbook);
        }
    }

    private static boolean isMergedRegionAlreadyExists(Sheet sheet, CellRangeAddress range) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress existingRegion = sheet.getMergedRegion(i);
            if (existingRegion.getFirstRow() == range.getFirstRow()
                    && existingRegion.getLastRow() == range.getLastRow()
                    && existingRegion.getFirstColumn() == range.getFirstColumn()
                    && existingRegion.getLastColumn() == range.getLastColumn()) {
                return true;
            }
        }
        return false;
    }


    private static void setCellStyleAndAlignment(Sheet sheet, CellRangeAddress range, Workbook workbook) {
        Row row = sheet.getRow(range.getFirstRow());
        if (row == null) return;

        Cell firstCell = row.getCell(range.getFirstColumn());
        if (firstCell == null) {
            firstCell = row.createCell(range.getFirstColumn());
        }

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        firstCell.setCellStyle(cellStyle);

        // Lặp qua các ô trong vùng được gộp để gán cùng một CellStyle
        for (int i = range.getFirstColumn(); i <= range.getLastColumn(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            cell.setCellStyle(cellStyle);
        }
    }


    private static void copyRow(Row templateRow, Row newRow, Workbook workbook, Map<CellStyle, CellStyle> styleCache) {
        for (int i = 0; i < templateRow.getLastCellNum(); i++) {
            Cell templateCell = templateRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            if (templateCell != null) {
                copyCell(templateCell, newCell, workbook, styleCache);
            }
        }

        // Sao chép chiều cao hàng
        newRow.setHeight(templateRow.getHeight());
    }

    private static void copyCell(Cell templateCell, Cell newCell, Workbook workbook, Map<CellStyle, CellStyle> styleCache) {
        // Lấy hoặc tạo mới CellStyle
        CellStyle templateStyle = templateCell.getCellStyle();
        if (!styleCache.containsKey(templateStyle)) {
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(templateStyle);
            styleCache.put(templateStyle, newStyle);
        }
        newCell.setCellStyle(styleCache.get(templateStyle));

        // Copy giá trị
        switch (templateCell.getCellType()) {
            case STRING:
                newCell.setCellValue(templateCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(templateCell)) {
                    newCell.setCellValue(templateCell.getDateCellValue());
                } else {
                    newCell.setCellValue(templateCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                newCell.setCellValue(templateCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(templateCell.getCellFormula());
                break;
            case BLANK:
                newCell.setBlank();
                break;
            default:
                break;
        }

        // Copy comment nếu có
        if (templateCell.getCellComment() != null) {
            newCell.setCellComment(templateCell.getCellComment());
        }
    }

    private static void updateTemplateWithCustomerData(Sheet sheet, Customer customer) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    cell.setCellValue(replacePlaceholders(cellValue, customer));
                }
            }
        }
    }

    // Hàm thay thế placeholders
    private static String replacePlaceholders(String cellValue, Customer customer) {
        return cellValue
                .replace("{{id}}", String.valueOf(customer.getCustomerId()))
                .replace("{{fullName}}", customer.getFullName())
                .replace("{{oldIndex}}", String.valueOf(customer.getOldIndex()))
                .replace("{{newIndex}}", String.valueOf(customer.getNewIndex()))
                .replace("{{unitPrice}}", String.valueOf(customer.getUnitPrice()));
    }

}

