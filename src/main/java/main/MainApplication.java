package main;

import process.ExcelProcessor;

public class MainApplication {
    public static final String ROOT_SOURCE = "D:\\Loclt\\Project_out\\loclt7\\DaElectric\\src\\";

    public static void main(String[] args) {
        String managementFilePath = ROOT_SOURCE + "ElectricityMoney_file.xlsx";
        String templateFilePath = ROOT_SOURCE + "HoaDon2023_Template.xlsx";
        String outputFilePath = ROOT_SOURCE + "HoaDonKetQua.xlsx";

        try {
            ExcelProcessor processor = new ExcelProcessor();
            processor.processFiles(managementFilePath, templateFilePath, outputFilePath);
            System.out.println("File processed successfully! Output saved to: " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
