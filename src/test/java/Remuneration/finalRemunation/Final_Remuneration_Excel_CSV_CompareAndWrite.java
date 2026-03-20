package Remuneration.finalRemunation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Final_Remuneration_Excel_CSV_CompareAndWrite {

    public static void main(String[] args) throws Exception {

        String filePath = "F:\\GCC-TBC October  repeaters\\Book1.xlsx";

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet originalSheet = workbook.getSheet("Original Copy");
        Sheet converterSheet = workbook.getSheet("Converter");

        List<String[]> mismatchRows = new ArrayList<>();

        // Convert Converter sheet to list
        List<Row> converterRows = new ArrayList<>();
        for (int i = 1; i <= converterSheet.getLastRowNum(); i++) {
            Row r = converterSheet.getRow(i);
            if (r != null) converterRows.add(r);
        }

        // Iterate Original Sheet
        for (int i = 1; i <= originalSheet.getLastRowNum(); i++) {

            Row originalRow = originalSheet.getRow(i);
            if (originalRow == null) continue;

            String ifscO = getCellValue(originalRow, 13);
            String accO  = getCellValue(originalRow, 14);
            String nameO = getCellValue(originalRow, 15);
            String amtO  = getCellValue(originalRow, 17);

            // Store converter match if found
            String ifscC = "", accC = "", nameC = "", amtC = "";
            boolean foundMatch = false;

            // Search for match
            for (Row convRow : converterRows) {

                String cIFSC = getCellValue(convRow, 2);
                String cAcc  = getCellValue(convRow, 3);
                String cName = getCellValue(convRow, 4);
                String cAmt  = getCellValue(convRow, 5);

                if (ifscO.equals(cIFSC) &&
                        accO.equals(cAcc) &&
                        nameO.equals(cName) &&
                        amtO.equals(cAmt)) {

                    foundMatch = true;

                    ifscC = cIFSC;
                    accC  = cAcc;
                    nameC = cName;
                    amtC  = cAmt;

                    break;
                }
            }

            // Add to mismatch Excel (even if match found, we show side-by-side once)
            if (!foundMatch) {
                mismatchRows.add(new String[]{
                        String.valueOf(i + 1),
                        ifscO, accO, nameO, amtO,
                        "", "", "", ""
                });
            }
//            else {
//                mismatchRows.add(new String[]{
//                        String.valueOf(i + 1),
//                        ifscO, accO, nameO, amtO,
//                        ifscC, accC, nameC, amtC
//                });
//            }
        }

        workbook.close();
        fis.close();

        writeMismatchExcel(mismatchRows);
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

    // -------------------- WRITE SIDE-BY-SIDE OUTPUT --------------------
    private static void writeMismatchExcel(List<String[]> data) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("SideBySide");

        // Header
        Row header = sheet.createRow(0);

        String[] titles = {
                "Original Row No",
                "IFSC (O)", "Account (O)", "Name (O)", "Amount (O)",
                "IFSC (C)", "Account (C)", "Name (C)", "Amount (C)"
        };

        for (int i = 0; i < titles.length; i++) {
            header.createCell(i).setCellValue(titles[i]);
        }

        // Add rows
        int rowIndex = 1;
        for (String[] rowData : data) {
            Row row = sheet.createRow(rowIndex++);
            for (int col = 0; col < rowData.length; col++) {
                row.createCell(col).setCellValue(rowData[col]);
            }
        }

        // Auto-size columns
        for (int i = 0; i < titles.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Save file
        FileOutputStream fos = new FileOutputStream("F:\\GCC-TBC October  repeaters\\SideBySide_MismatchFinal.xlsx");
        wb.write(fos);
        fos.close();
        wb.close();

        System.out.println("Side-by-Side Excel created successfully!");
    }
}
