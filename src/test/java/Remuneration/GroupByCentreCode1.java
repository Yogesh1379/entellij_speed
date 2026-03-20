package Remuneration;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class GroupByCentreCode1 {
        static List<String> mismatchList = new ArrayList<>();
    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(
                new File("F:\\GCC-TBC October  repeaters\\22-11-2025_NEW.xlsx")
        );
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("Sheet1");
        int centreCodeCol = 1;
        Map<String, List<Row>> map = new HashMap<>();
        // STEP 1 → Group rows by Centre Code
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String centreCode = getValue(row.getCell(centreCodeCol));
            if (centreCode.isEmpty()) continue;
            map.computeIfAbsent(centreCode, k -> new ArrayList<>()).add(row);
        }
        // STEP 2 → Process each centre
        for (String centreCode : map.keySet()) {
            List<Row> rows = map.get(centreCode);
            validateCentre(centreCode, rows);
        }
        workbook.close();
        fis.close();
        // STEP 3 → Write mismatches into new Excel file
        writeMismatchExcel();
        System.out.println("\n✔ Mismatch Excel Successfully Generated!");
    }

    // ===================== VALIDATION  =====================
    static void validateCentre(String centreCode, List<Row> rows) {

        int centreNameCol = 2;
        int engDaysCol = 3;
        int engPcCol = 4;
        int marDaysCol = 6;
        int marPcCol = 7;
        int engMarDaysCol = 10;
        int totalDaysCol = 12;
        int designationCol = 16;
        int totalAmtCol = 17;
        // Condition 1: All rows must have same centre name
        String firstName = getValue(rows.get(0).getCell(centreNameCol));
        for (Row r : rows) {
            String name = getValue(r.getCell(centreNameCol));
            if (!name.equals(firstName)) {
                addMismatch(centreCode, "Name mismatch | Row " + r.getRowNum());
            }
        }
        // Condition 2 – manpower /amount validation
        String engPCct = getValue(rows.get(0).getCell(engPcCol));
        String marPCct = getValue(rows.get(0).getCell(marPcCol));
        int EngPCCount = parseInt(engPCct);
        int MarPCCount = parseInt(marPCct);
        int maxPCCount = Math.max(EngPCCount, MarPCCount);
        int requiredmanPower = (int) Math.ceil(maxPCCount / 30.0);
        // --- IT Teacher ---
        int itTeacherCount = 0;
        for (Row r : rows) {
            if (getValue(r.getCell(designationCol)).contains("IT Teacher"))
                itTeacherCount++;
        }
        if (itTeacherCount != requiredmanPower) {
            addMismatch(centreCode, "IT Teacher manpower mismatch | Required=" + requiredmanPower + " Provided=" + itTeacherCount);
        }

        // Check amount
        for (Row r : rows) {
            if (getValue(r.getCell(designationCol)).contains("IT Teacher")) {

                int totalDays1 = parseInt(getValue(r.getCell(totalDaysCol)));
                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays1 * 600;

                if (rowAmount != expected) {
                    addMismatch(centreCode, "Incorrect IT Teacher Amount | Row " + r.getRowNum()
                            + " | Expected=" + expected + " Found=" + rowAmount);
                }
            }
        }

        // --- Winner Technical ---
        int winnerTechnical = 0;
        for (Row r : rows) {
            if (getValue(r.getCell(designationCol)).contains("Winner Technical"))
                winnerTechnical++;
        }
        if (winnerTechnical != requiredmanPower) {
            addMismatch(centreCode, "Winner Technical manpower mismatch | Required="
                    + requiredmanPower + " Provided=" + winnerTechnical);
        }

        // Winner technical rate check
        for (Row r : rows) {
            if (getValue(r.getCell(designationCol)).contains("Winner Technical")) {

                int totalDays1 = parseInt(getValue(r.getCell(totalDaysCol)));
                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays1 * 400;

                if (rowAmount != expected) {
                    addMismatch(centreCode, "Incorrect Winner Technical Amount | Row " + r.getRowNum()
                            + " | Expected=" + expected + " Found=" + rowAmount);
                }
            }
        }

        // --- Peon ---
        int peon = 0;
        for (Row r : rows) {
            if (getValue(r.getCell(designationCol)).contains("Peon"))
                peon++;
        }

        if (peon != requiredmanPower) {
            addMismatch(centreCode, "Peon manpower mismatch | Required=" + requiredmanPower + " Provided=" + peon);
        }

        // check peon rate
        for (Row r : rows) {
            if (getValue(r.getCell(designationCol)).contains("Peon")) {

                int engday = parseInt(getValue(r.getCell(engDaysCol)));
                int marday = parseInt(getValue(r.getCell(marDaysCol)));
                int totalDays = engday + marday;

                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays * 150;

                if (rowAmount != expected) {
                    addMismatch(centreCode, "Incorrect Peon Amount | Row " + r.getRowNum()
                            + " | Expected=" + expected + " Found=" + rowAmount);
                }
            }
        }
    }

    // Add mismatch to list
    static void addMismatch(String centreCode, String msg) {
        mismatchList.add(centreCode + " | " + msg);
    }

    // ==================== WRITE TO EXCEL ======================
    static void writeMismatchExcel() throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Mismatches");

        int rowIndex = 0;

        // Header
        Row header = sheet.createRow(rowIndex++);
        header.createCell(0).setCellValue("Centre Code");
        header.createCell(1).setCellValue("Mismatch Description");

        for (String entry : mismatchList) {

            String[] parts = entry.split("\\|", 2);

            Row row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(parts[0].trim());
            row.createCell(1).setCellValue(parts.length > 1 ? parts[1].trim() : "");
        }

        FileOutputStream fos = new FileOutputStream("F:\\GCC-TBC October  repeaters\\Mismatch_Report.xlsx");
        wb.write(fos);
        fos.close();
        wb.close();
    }

    // Helper Functions
    private static String getValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
        }
        return "";
    }

    private static int parseInt(String s) {
        if (s == null || s.trim().isEmpty()) return 0;
        try { return Integer.parseInt(s.trim()); }
        catch (Exception e) { return 0; }
    }
}
