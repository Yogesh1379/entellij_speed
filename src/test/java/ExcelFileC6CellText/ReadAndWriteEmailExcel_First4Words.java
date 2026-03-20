package ExcelFileC6CellText;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadAndWriteEmailExcel_First4Words {

    public static void main(String[] args) {
        String inputPath = "C:\\Users\\User\\Desktop\\Student EmailAnswer.xlsx";   // existing Excel file
        String outputPath = "C:\\Users\\User\\Desktop\\EmailContentOutput.xlsx"; // new Excel file

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook inputWorkbook = new XSSFWorkbook(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = inputWorkbook.getSheetAt(0); // read first sheet
            Sheet outputSheet = outputWorkbook.createSheet("Result");

            // Header Row
            Row header = outputSheet.createRow(0);
            header.createCell(0).setCellValue("Seat Number");
            header.createCell(1).setCellValue("Content");
            header.createCell(2).setCellValue("Batch");
            header.createCell(3).setCellValue("subject");

            int outRowNum = 1;
            DataFormatter df = new DataFormatter();

            // Iterate all rows (starting from row 1 to skip header)
            for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
                Row row = inputSheet.getRow(i);
                if (row == null) continue;

                // Column F → index 5 (A=0, B=1, ..., F=5)
                Cell textCell = row.getCell(5);
                // Column AW → index 48 (A=0, ..., Z=25, AA=26, ..., AW=48)
                Cell seatCell = row.getCell(48);
                Cell batchCell = row.getCell(22);
                Cell subjectCell = row.getCell(6);

                String fullText = df.formatCellValue(textCell).trim();
                String seatNumber = df.formatCellValue(seatCell).trim();
                String Batch = df.formatCellValue(batchCell).trim();
                String subject = df.formatCellValue(subjectCell).trim();

                if (fullText.isEmpty() && seatNumber.isEmpty()) continue;

                String first4Words = extractFirst4Words(fullText);

                Row outRow = outputSheet.createRow(outRowNum++);
                outRow.createCell(0).setCellValue(seatNumber);
                outRow.createCell(1).setCellValue(first4Words);
                outRow.createCell(2).setCellValue(Batch);
                outRow.createCell(3).setCellValue(subject);
            }

            // Write to new Excel
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                outputWorkbook.write(fos);
            }

            System.out.println("✅ Data written successfully to: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Helper method to extract first 4 words
    private static String extractFirst4Words(String text) {
        if (text == null || text.isEmpty()) return "";
        String[] words = text.split("\\s+");
        int limit = Math.min(words.length, 4);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < limit; i++) {
            sb.append(words[i]);
            if (i < limit - 1) sb.append(" ");
        }
        return sb.toString();
    }
}
