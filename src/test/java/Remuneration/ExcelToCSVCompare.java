package Remuneration;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelToCSVCompare {

    public static void main(String[] args) throws Exception {

        String filePath = "F:\\GCC-TBC October  repeaters\\Book1.xlsx";

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet originalSheet = workbook.getSheet("Original Copy");
        Sheet converterSheet = workbook.getSheet("Converter");

        List<String> mismatches = new ArrayList<>();

        // Convert converter sheet into list of rows for faster searching
        List<Row> converterRows = new ArrayList<>();
        for (int i = 1; i <= converterSheet.getLastRowNum(); i++) {
            Row r = converterSheet.getRow(i);
            if (r != null) converterRows.add(r);
        }

        // Iterate entire Original sheet
        for (int i = 1; i <= originalSheet.getLastRowNum(); i++) {

            Row originalRow = originalSheet.getRow(i);
            if (originalRow == null) continue;

            String ifscO = getCellValue(originalRow, 13);
            String accO  = getCellValue(originalRow, 14);
            String nameO = getCellValue(originalRow, 15);
            String amtO  = getCellValue(originalRow, 17);

            boolean foundMatch = false;

            // Search entire Converter sheet
            for (Row convRow : converterRows) {

                String ifscC = getCellValue(convRow, 2);
                String accC  = getCellValue(convRow, 3);
                String nameC = getCellValue(convRow, 4);
                String amtC  = getCellValue(convRow, 5);

                if (ifscO.equals(ifscC) &&
                        accO.equals(accC) &&
                        nameO.equals(nameC) &&
                        amtO.equals(amtC)) {

                    foundMatch = true;
                    break;
                }
            }

            // If no matching record found in converter sheet
            if (!foundMatch) {
                mismatches.add("No match found for Original Row " + (i + 1)
                        + " | IFSC=" + ifscO
                        + ", Account=" + accO
                        + ", Name=" + nameO
                        + ", Amount=" + amtO);
            }
        }

        workbook.close();
        fis.close();

        // Print mismatches
        if (mismatches.isEmpty()) {
            System.out.println("All records in Original Copy matched with Converter!");
        } else {
            mismatches.forEach(System.out::println);
        }
    }

    private static String getCellValue(Row row, int colIndex) {
        if (row == null) return "";
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
            case FORMULA -> cell.getCellFormula();
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }
}
