package example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CombineAndCenterContentExample {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // ... (code tạo các cell và điền dữ liệu như trên)
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("Cell 1");
        Cell cell2 = row.createCell(1);
        cell2.setCellValue("Cell 2");

        // Gộp và căn giữa
        CellRangeAddress mergedRegion = new CellRangeAddress(0, 0, 0, 1);
        sheet.addMergedRegion(mergedRegion);


        Cell mergedCell = row.getCell(0);
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        mergedCell.setCellStyle(style);

        // Ghi workbook vào file
        FileOutputStream outputStream = new FileOutputStream("merged_cells.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }


}
