package example;

import model.Customer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import static example.ReadExcelExample.*;

public class CloneTemplate {
    public static void main(String[] args) {

        try {
            List<Customer> customers = readCustomersFromExcel(filePath);

            // Mở file template
            Workbook workbook = new XSSFWorkbook(new FileInputStream(invoiceTemplatePath));
            Sheet templateSheet = workbook.getSheetAt(0);

            for (Customer customer : customers) {
                // Update dữ liệu trong sheet template
                 updateTemplateWithCustomerData(templateSheet, customer);
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

    /*private static void updateTemplateWithCustomerData(Sheet sheet, Customer customer) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    // Kiểm tra và thay thế placeholder
                    if (cellValue.contains("{{id}}")) {
                        cell.setCellValue(cellValue.replace("{{id}}", String.valueOf(customer.getCustomerId())));
                    }
                    if (cellValue.contains("{{fullName}}")) {
                        cell.setCellValue(cellValue.replace("{{fullName}}", customer.getFullName()));
                    }
                    if (cellValue.contains("{{oldIndex}}")) {
                        try {
                            double oldIndex = Double.parseDouble(String.valueOf(customer.getOldIndex()));
                            cell.setCellValue(oldIndex);
                        } catch (NumberFormatException e) {
                            cell.setCellValue(cellValue.replace("{{oldIndex}}", String.valueOf(customer.getOldIndex())));
                        }
                    }
                    if (cellValue.contains("{{newIndex}}")) {
                        try {
                            double newIndex = Double.parseDouble(String.valueOf(customer.getNewIndex()));
                            cell.setCellValue(newIndex);
                        } catch (NumberFormatException e) {
                            cell.setCellValue(cellValue.replace("{{newIndex}}", String.valueOf(customer.getNewIndex())));
                        }
                    }
                    if (cellValue.contains("{{unitPrice}}")) {
                        cell.setCellValue(customer.getUnitPrice());
                    }
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    String cellValue = cell.toString();

                    // Kiểm tra và thay thế với số
                    if (cellValue.contains("{{id}}")) {
                        cell.setCellValue(customer.getCustomerId());
                    }
                    if (cellValue.contains("{{oldIndex}}")) {
                        cell.setCellValue(customer.getOldIndex());
                    }
                    if (cellValue.contains("{{newIndex}}")) {
                        cell.setCellValue(customer.getNewIndex());
                    }
                    if (cellValue.contains("{{unitPrice}}")) {
                        cell.setCellValue(customer.getUnitPrice());
                    }
                }
            }
        }
    }*/

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

