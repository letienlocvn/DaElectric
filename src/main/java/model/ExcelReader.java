package model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    /**
     * <ul>
     *  <li> STT (Index) </li>
     *  <li> Họ Tên (Customer Name)</li>
     *  <li> Chỉ số cũ (Old Index)</li>
     *  <li> Chỉ số mới (New Index)</li>
     *  <li> Số trong tháng (Units in Month)</li>
     *  <li> Đơn giá (Unit Price)</li>
     *  <li> Thành tiền (Amount)</li>
     *  <li> Công ghi điện (Recording Fee)</li>
     *  <li> Tổng thanh toán (Total Payment)</li>
     * </ul>
     */
    public static class ElectricBillRecord {
        public int index;
        public String customerName;
        public double oldIndex;
        public double newIndex;
        public double unitsInMonth;
        public double unitPrice;
        public double amount;
        public double recordingFee;
        public double totalPayment;

        @Override
        public String toString() {
            return "ElectricBillRecord{" +
                    "index=" + index +
                    ", customerName='" + customerName + '\'' +
                    ", oldIndex=" + oldIndex +
                    ", newIndex=" + newIndex +
                    ", totalPayment=" + totalPayment +
                    '}';
        }
    }

    public static List<ElectricBillRecord> readInputFile(String filePath) throws IOException {
        List<ElectricBillRecord> records = new ArrayList<>();
        try (
                FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                // Skip header row
                Row row = sheet.getRow(i + 1);
                if (row == null) continue;
                Cell indexCell = row.getCell(0);
                if (indexCell == null || indexCell.getCellType() != CellType.NUMERIC) {
                    continue; // Skip rows where the index cell is missing or not numeric
                }
                ElectricBillRecord record = new ElectricBillRecord();
                record.index = (int) row.getCell(0).getNumericCellValue();
                record.customerName = row.getCell(1).getStringCellValue();
                record.oldIndex = row.getCell(2).getNumericCellValue();
                record.newIndex = row.getCell(3).getNumericCellValue();
                record.unitsInMonth = row.getCell(4).getNumericCellValue();
                record.unitPrice = row.getCell(5).getNumericCellValue();
                record.amount = row.getCell(6).getNumericCellValue();
                record.recordingFee = row.getCell(7).getNumericCellValue();
                record.totalPayment = row.getCell(8).getNumericCellValue();

                records.add(record);
            }
        }

        return records;
    }
}
