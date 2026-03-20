package Remuneration;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class remuneration1 {

    static class CollegeRecord {
        String centreCode;
        String centreName;
        String englishExamDays;
        String englishPCCount;
        String marathiExamDays;
        String marathiPCCount;
        String englishMarathiExamDays;
        String totalDaysWithMockTest;

        @Override
        public String toString() {
            return centreCode + " | " + centreName + " | " + englishExamDays + " | "
                    + englishPCCount + " | " + marathiExamDays + " | " + marathiPCCount
                    + " | " + englishMarathiExamDays + " | " + totalDaysWithMockTest;
        }
    }

    public static void main(String[] args) throws IOException {
        File inputfile = new File("F:\\GCC-TBC October  repeaters\\All1_Region_Remuneration _OCT_Updated21112025.xlsx");
        FileInputStream fis = new FileInputStream(inputfile);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        // Sheets
        Sheet sheetCollege = workbook.getSheet("college");
        Sheet sheetITTeacher = workbook.getSheet("It_Teacher");
        Sheet sheetExtraIT = workbook.getSheet("Extra IT Teacher");
        Sheet sheetWinnerTech = workbook.getSheet("Extra Winner Technical");

        List<CollegeRecord> list = new ArrayList<>();

        // Read College Sheet
        int centreCodeCol = 1, centreNameCol = 2, engExamDaysCol = 3, engPCCountCol = 4,
                marathiExamDaysCol = 6, marathiPCCountCol = 7, engMarathiExamDaysCol = 10,
                totalDaysExamCol = 12;

        for (int r = 1; r <= sheetCollege.getLastRowNum(); r++) {
            Row row = sheetCollege.getRow(r);
            if (row == null) continue;

            CollegeRecord rec = new CollegeRecord();
            rec.centreCode = getValue(row.getCell(centreCodeCol));
            rec.centreName = getValue(row.getCell(centreNameCol));
            rec.englishExamDays = getValue(row.getCell(engExamDaysCol));
            rec.englishPCCount = getValue(row.getCell(engPCCountCol));
            rec.marathiExamDays = getValue(row.getCell(marathiExamDaysCol));
            rec.marathiPCCount = getValue(row.getCell(marathiPCCountCol));
            rec.englishMarathiExamDays = getValue(row.getCell(engMarathiExamDaysCol));
            rec.totalDaysWithMockTest = getValue(row.getCell(totalDaysExamCol));
            list.add(rec);
        }

        // Process IT Teacher Sheet (without 30-PC logic)
        processRegularSheet(sheetITTeacher, list, 600);

        // Process Extra IT Teacher (with 30-PC rule)
        processExtraSheet(sheetExtraIT, list, 600);

        // Process Winner Technical Sheet (without 30-PC logic)
        processRegularSheet(sheetWinnerTech, list, 400);

        // Process Extra Winner Technical (with 30-PC rule)
        processExtraSheet(sheetWinnerTech, list, 400);

        // Close input
        fis.close();

        // Write workbook
        FileOutputStream fos = new FileOutputStream("F:\\GCC-TBC October  repeaters\\Remuneration_Updated_Result12.xlsx");
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Comparison complete. Results written to Excel.");
    }

    private static void processRegularSheet(Sheet sheet, List<CollegeRecord> list, int perDayRate) {
        int centreCodeCol = 1, centreNameCol = 2, engExamDaysCol = 3, engPCCountCol = 4,
                marathiExamDaysCol = 6, marathiPCCountCol = 7, engMarathiExamDaysCol = 10,
                totalDaysExamCol = 12, totalAmountCol = 17, resultCol = 20, mismatchCol = 21,
                wrongAmountCol = 22;

        // Headers
        Row header = sheet.getRow(0);
        if (header == null) header = sheet.createRow(0);
        header.createCell(resultCol).setCellValue("Result");
        header.createCell(mismatchCol).setCellValue("Mismatch Details");
        header.createCell(wrongAmountCol).setCellValue("Amount");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String cCode = getValue(row.getCell(centreCodeCol));
            String cName = getValue(row.getCell(centreNameCol));
            String engDays = getValue(row.getCell(engExamDaysCol));
            String engPc = getValue(row.getCell(engPCCountCol));
            String marDays = getValue(row.getCell(marathiExamDaysCol));
            String marPc = getValue(row.getCell(marathiPCCountCol));
            String engMarDays = getValue(row.getCell(engMarathiExamDaysCol));
            String totalDays = getValue(row.getCell(totalDaysExamCol));
            String totalAmountExcel = getValue(row.getCell(totalAmountCol));

            int exlAmt = parseIntSafe(totalAmountExcel);
            int ttlDays = parseIntSafe(totalDays);

            Cell amtCell = row.getCell(wrongAmountCol);
            if (amtCell == null) amtCell = row.createCell(wrongAmountCol);
            int expectedAmount = ttlDays * perDayRate;
            amtCell.setCellValue(exlAmt != expectedAmount ? expectedAmount : 0);


            boolean found = false;
            String result = "";
            String mismatchDetails = "";
            List<String> mismatches = new ArrayList<>();

            for (CollegeRecord rec : list) {
                if (cCode.equals(rec.centreCode)) {
                    found = true;
                    if (!cName.equals(rec.centreName)) mismatches.add("CentreName");
                    if (!engDays.equals(rec.englishExamDays)) mismatches.add("EnglishExamDays");
                    if (!engPc.equals(rec.englishPCCount)) mismatches.add("EnglishPCCount");
                    if (!marDays.equals(rec.marathiExamDays)) mismatches.add("MarathiExamDays");
                    if (!marPc.equals(rec.marathiPCCount)) mismatches.add("MarathiPCCount");
                    if (!engMarDays.equals(rec.englishMarathiExamDays)) mismatches.add("EnglishMarathiExamDays");
                    if (!totalDays.equals(rec.totalDaysWithMockTest)) mismatches.add("TotalDaysWithMockTest");
                    break;
                }
            }

            if (!found) result = "Not Found";
            else if (mismatches.isEmpty()) result = "Match";
            else {
                result = "Mismatch";
                mismatchDetails = String.join(", ", mismatches);
            }

            Cell resCell = row.getCell(resultCol);
            if (resCell == null) resCell = row.createCell(resultCol);
            resCell.setCellValue(result);

            Cell misCell = row.getCell(mismatchCol);
            if (misCell == null) misCell = row.createCell(mismatchCol);
            misCell.setCellValue(mismatchDetails);
        }
    }

    private static void processExtraSheet(Sheet sheet, List<CollegeRecord> list, int perDayRate) {
        int centreCodeCol = 1, centreNameCol = 2, engExamDaysCol = 3, engPCCountCol = 4,
                marathiExamDaysCol = 6, marathiPCCountCol = 7, engMarathiExamDaysCol = 10,
                totalDaysExamCol = 12, totalAmountCol = 17, resultCol = 20, mismatchCol = 21,
                wrongAmountCol = 22;

        Row header = sheet.getRow(0);
        if (header == null) header = sheet.createRow(0);
        header.createCell(resultCol).setCellValue("Result");
        header.createCell(mismatchCol).setCellValue("Mismatch Details");
        header.createCell(wrongAmountCol).setCellValue("Amount");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String cCode = getValue(row.getCell(centreCodeCol));
            String cName = getValue(row.getCell(centreNameCol));
            String engDays = getValue(row.getCell(engExamDaysCol));
            String engPc = getValue(row.getCell(engPCCountCol));
            String marDays = getValue(row.getCell(marathiExamDaysCol));
            String marPc = getValue(row.getCell(marathiPCCountCol));
            String engMarDays = getValue(row.getCell(engMarathiExamDaysCol));
            String totalDays = getValue(row.getCell(totalDaysExamCol));
            String totalAmountExcel = getValue(row.getCell(totalAmountCol));

            // 30-PC RULE
            if ("30".equals(engPc) || "30".equals(marPc)) {
                engDays = "1";
                marDays = "1";
                engMarDays = "2";
                totalDays = "2";
                row.getCell(engExamDaysCol).setCellValue("1");
                row.getCell(marathiExamDaysCol).setCellValue("1");
                row.getCell(engMarathiExamDaysCol).setCellValue("2");
                row.getCell(totalDaysExamCol).setCellValue("2");
            }

            int exlAmt = parseIntSafe(totalAmountExcel);
            int ttlDays = parseIntSafe(totalDays);

            Cell amtCell = row.getCell(wrongAmountCol);
            if (amtCell == null) amtCell = row.createCell(wrongAmountCol);
            int expectedAmount = ttlDays * perDayRate;
            amtCell.setCellValue(exlAmt != expectedAmount ? expectedAmount : 0);


            boolean found = false;
            String result = "";
            String mismatchDetails = "";
            List<String> mismatches = new ArrayList<>();

            for (CollegeRecord rec : list) {
                if (cCode.equals(rec.centreCode)) {
                    found = true;
                    if (!cName.equals(rec.centreName)) mismatches.add("CentreName");
                    if (!engDays.equals(rec.englishExamDays)) mismatches.add("EnglishExamDays");
                    if (!engPc.equals(rec.englishPCCount)) mismatches.add("EnglishPCCount");
                    if (!marDays.equals(rec.marathiExamDays)) mismatches.add("MarathiExamDays");
                    if (!marPc.equals(rec.marathiPCCount)) mismatches.add("MarathiPCCount");
                    if (!engMarDays.equals(rec.englishMarathiExamDays)) mismatches.add("EnglishMarathiExamDays");
                    if (!totalDays.equals(rec.totalDaysWithMockTest)) mismatches.add("TotalDaysWithMockTest");
                    break;
                }
            }

            if (!found) result = "Not Found";
            else if (mismatches.isEmpty()) result = "Match";
            else {
                result = "Mismatch";
                mismatchDetails = String.join(", ", mismatches);
            }

            // Group check for extra sheets
            int pcEng = parseIntSafe(engPc);
            int pcMar = parseIntSafe(marPc);
            int maxPC = Math.max(pcEng, pcMar);
            int groups = (int) Math.ceil(maxPC / 30.0);

            for (int g = 1; g <= groups-1; g++) {
                String groupCode = cCode + "-" + g;
                boolean exists = searchGroup(sheet, groupCode, cName);
                if (!exists) {
                    mismatchDetails += " | Missing Group: " + groupCode;
                    result = "Mismatch";
                }
            }

            Cell resCell = row.getCell(resultCol);
            if (resCell == null) resCell = row.createCell(resultCol);
            resCell.setCellValue(result);

            Cell misCell = row.getCell(mismatchCol);
            if (misCell == null) misCell = row.createCell(mismatchCol);
            misCell.setCellValue(mismatchDetails);
        }
    }

    private static boolean searchGroup(Sheet sheet, String code, String name) {
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String cCode = getValue(row.getCell(1));
            String cName = getValue(row.getCell(2));
            if (code.equals(cCode) && name.equals(cName)) return true;
        }
        return false;
    }

    private static String getValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue().toString();
                return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }

    private static int parseIntSafe(String s) {
        try { return (s == null || s.trim().isEmpty()) ? 0 : Integer.parseInt(s.trim()); }
        catch (Exception e) { return 0; }
    }
}
