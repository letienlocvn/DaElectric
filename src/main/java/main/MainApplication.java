package main;

import model.ExcelReader;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import static model.ExcelReader.readInputFile;
import static process.InvoiceGenerator.mapDataToTemplate;

public class MainApplication {
    public static final String INPUT = "data/final/input";
    public static final String OUTPUT = "data/final/output";

    public static void main(String[] args) throws IOException {
        List<ExcelReader.ElectricBillRecord> records =
                readInputFile(INPUT + "/ElectricityManagement.xlsx");
        records.forEach(System.out::println);

        // Example data (replace with actual input)
        List<Map<String, Object>> data = List.of(
                Map.of("index", "1", "fullName", "John Doe", "oldIndex", "100", "newIndex", "200", "unitPrice", "5000"),
                Map.of("index", "2", "fullName", "Jane Smith", "oldIndex", "150", "newIndex", "250", "unitPrice", "4500")
        );

        String templatePath = "HoaDon2023_Template.xlsx";
        String outputPath = "GeneratedInvoices.xlsx";

        mapDataToTemplate(templatePath, outputPath, data);

        System.out.println("Invoices generated successfully!");
    }
}
