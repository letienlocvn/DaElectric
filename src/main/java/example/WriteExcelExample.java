package example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelExample {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Example");
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        cell.setCellValue("Le Tien Loc");

        try (FileOutputStream stream = new FileOutputStream("example.xlsx")){
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
