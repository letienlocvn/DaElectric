package example;

import model.Customer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import static configure.FileConfiguration.*;

public class ReadExcelExample {


    public static void main(String[] args) {
        try {
            List<Customer> customers = readCustomersFromExcel(filePath);
            Workbook invoiceTemplate = new XSSFWorkbook(new FileInputStream(invoiceTemplatePath));

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Invoices");

            // Append invoices for each customer
            int currentRow = 0;
            for (Customer customer : customers) {
                currentRow = appendInvoiceForCustomer(invoiceTemplate, outputSheet, customer, currentRow);
            }

            // Save to output file
            FileOutputStream fileOut = new FileOutputStream(outputPath);
            outputWorkbook.write(fileOut);
            fileOut.close();
            invoiceTemplate.close();
            outputWorkbook.close();

            System.out.println("Invoices generated successfully!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static int appendInvoiceForCustomer(Workbook templateWorkbook, Sheet outputSheet, Customer customer, int currentRow) {
        Sheet templateSheet = templateWorkbook.getSheetAt(0);

        for (int rowIndex = 0; rowIndex <= templateSheet.getLastRowNum(); rowIndex++) {
            Row templateRow = templateSheet.getRow(rowIndex);
            Row outputRow = outputSheet.createRow(currentRow++);

            if (templateRow != null) {
                // Sao chép chiều cao hàng từ template
                outputRow.setHeight(templateRow.getHeight());

                for (int cellIndex = 0; cellIndex < templateRow.getLastCellNum(); cellIndex++) {
                    Cell templateCell = templateRow.getCell(cellIndex);
                    Cell outputCell = outputRow.createCell(cellIndex);

                    if (templateCell != null) {
                        // Sao chép định dạng từ template
                        copyCellStyle(templateCell, outputCell);

                        String cellValue = templateCell.toString();
                        // Replace placeholders with customer data
                        cellValue = cellValue.replace("{{id}}", String.valueOf(customer.getCustomerId()))
                                .replace("{{fullName}}", customer.getFullName())
                                .replace("{{oldIndex}}", String.valueOf(customer.getOldIndex()))
                                .replace("{{newIndex}}", String.valueOf(customer.getNewIndex()))
                                .replace("{{unitPrice}}", String.valueOf(customer.getUnitPrice()));

                        // Kiểm tra loại dữ liệu ban đầu
                        switch (templateCell.getCellType()) {
                            case STRING -> outputCell.setCellValue(cellValue);
                            case NUMERIC -> {
                                if (templateCell.getCellStyle().getDataFormatString().contains("0.00") ||
                                        templateCell.getCellStyle().getDataFormatString().contains("#")) {
                                    outputCell.setCellValue(Double.parseDouble(cellValue));
                                } else {
                                    outputCell.setCellValue(cellValue);
                                }
                            }
                            default -> outputCell.setCellValue(cellValue);
                        }
                    }
                }
            }
        }

        return currentRow;
    }

    private static void copyCellStyle(Cell templateCell, Cell outputCell) {
        Workbook workbook = outputCell.getSheet().getWorkbook();
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.cloneStyleFrom(templateCell.getCellStyle());
        outputCell.setCellStyle(newCellStyle);
    }

    public static List<Customer> readCustomersFromExcel(String filePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        List<Customer> customers = new ArrayList<>();
        for (int rowNum = 1; rowNum < sheet.getLastRowNum() + 1; rowNum++) {
            Row row = sheet.getRow(rowNum);

            // Create a new Customer object
            Customer customer = new Customer();

            // Set the customer's properties based on the cell values
            customer.setCustomerId((int) row.getCell(0).getNumericCellValue());
            customer.setFullName(row.getCell(1).getStringCellValue());
            customer.setOldIndex((int) row.getCell(2).getNumericCellValue());
            customer.setNewIndex((int) row.getCell(3).getNumericCellValue());
            customer.setConsumedElectricity((int) row.getCell(4).getNumericCellValue());
            customer.setUnitPrice(row.getCell(5).getNumericCellValue());
            customer.setElectricityCost(row.getCell(6).getNumericCellValue());
            customer.setMeterFee(row.getCell(7).getNumericCellValue());
            customer.setTotalPayment(row.getCell(8).getNumericCellValue());
            customer.setAddress(row.getCell(9).getStringCellValue());
            customer.setPhoneNumber(row.getCell(10).getStringCellValue());
            customer.setBillDate(row.getCell(11).getDateCellValue());

            // Add the customer to the list
            customers.add(customer);
        }

        // Close the workbook and input stream
        workbook.close();
        inputStream.close();

        return customers;
    }
}
