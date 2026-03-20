package Remuneration;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class remuneration {

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
        File inputfile = new File("F:\\GCC-TBC October  repeaters\\Remunaration.xlsx");
        FileInputStream fis = new FileInputStream(inputfile);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("college");

        List<CollegeRecord> list = new ArrayList<>();

        // Column indexes
        int centreCodeCol = 1;
        int centreNameCol = 2;
        int engExamDaysCol = 3;
        int engPCCountCol = 4;
        int marathiExamDaysCol = 6;
        int marathiPCCountCol = 7;
        int engMarathiExamDaysCol = 10;
        int totalDaysExamCol = 12;

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
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

        Sheet sheet1 = workbook.getSheet("It_Teacher");
        // Column indexes
        int centreCodeCol1 = 1;
        int centreNameCol1 = 2;
        int engExamDaysCol1 = 3;
        int engPCCountCol1 = 4;
        int marathiExamDaysCol1 = 6;
        int marathiPCCountCol1 = 7;
        int engMarathiExamDaysCol1 = 10;
        int totalDaysExamCol1 = 12;
        int Total_Amount=17;
        // Result column (you can choose an empty column, e.g., 20th column)
        int resultCol = 20;
        int mismatchCol = 21;
        int WrongAmount=22;
// Create headers
        Row headerRow = sheet1.getRow(0);
        if (headerRow == null) headerRow = sheet1.createRow(0);
        headerRow.createCell(resultCol).setCellValue("Result");
        headerRow.createCell(mismatchCol).setCellValue("Mismatch Details");
        headerRow.createCell(WrongAmount).setCellValue("Amount");

// Iterate rows
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row row1 = sheet1.getRow(i);
            if (row1 == null) continue;

            String cCode1 = getValue(row1.getCell(centreCodeCol1));
            String cName1 = getValue(row1.getCell(centreNameCol1));
            String engDays1 = getValue(row1.getCell(engExamDaysCol1));
            String engPcCnt1 = getValue(row1.getCell(engPCCountCol1));
            String marDays1 = getValue(row1.getCell(marathiExamDaysCol1));
            String marPcCnt1 = getValue(row1.getCell(marathiPCCountCol1));
            String engMarDays1 = getValue(row1.getCell(engMarathiExamDaysCol1));
            String totalDays1 = getValue(row1.getCell(totalDaysExamCol1));
            String TotalAmountExl=getValue(row1.getCell(Total_Amount));

            int ExlAmt = 0;
            int TtlDAys = 0;

            if (!TotalAmountExl.isEmpty()) {
                ExlAmt = Integer.parseInt(TotalAmountExl);
            }
            if (!totalDays1.isEmpty()) {
                TtlDAys = Integer.parseInt(totalDays1);
            }
            Cell amt = row1.getCell(WrongAmount);
            if (amt == null) amt = row1.createCell(WrongAmount); // Create if null

            if (ExlAmt != (TtlDAys * 600)) {
                amt.setCellValue(TtlDAys * 600);
            } else {
                amt.setCellValue(""); // Optional: clear if matches
            }
            boolean found = false;
            String result = "";
            String mismatchDetails = "";

            for (CollegeRecord rec : list) {
                if (cCode1.equals(rec.centreCode)) {
                    found = true;
                    List<String> mismatches = new ArrayList<>();

                    if (!cName1.equals(rec.centreName)) mismatches.add("CentreName");
                    if (!engDays1.equals(rec.englishExamDays)) mismatches.add("EnglishExamDays");
                    if (!engPcCnt1.equals(rec.englishPCCount)) mismatches.add("EnglishPCCount");
                    if (!marDays1.equals(rec.marathiExamDays)) mismatches.add("MarathiExamDays");
                    if (!marPcCnt1.equals(rec.marathiPCCount)) mismatches.add("MarathiPCCount");
                    if (!engMarDays1.equals(rec.englishMarathiExamDays)) mismatches.add("EnglishMarathiExamDays");
                    if (!totalDays1.equals(rec.totalDaysWithMockTest)) mismatches.add("TotalDaysWithMockTest");

                    if (mismatches.isEmpty()) {
                        result = "Match";
                    } else {
                        result = "Mismatch";
                        mismatchDetails = String.join(", ", mismatches);
                    }
                    break;
                }
            }

            if (!found) {
                result = "Not Found";
            }

            // Write to Excel
            Cell resultCell = row1.getCell(resultCol);
            if (resultCell == null) resultCell = row1.createCell(resultCol);
            resultCell.setCellValue(result);

            Cell mismatchCell = row1.getCell(mismatchCol);
            if (mismatchCell == null) mismatchCell = row1.createCell(mismatchCol);
            mismatchCell.setCellValue(mismatchDetails);
        }
//test
        Sheet sheet2 = workbook.getSheet("Extra IT Teacher");
        // Column indexes
        int centreCodeCol2 = 1;
        int centreNameCol2 = 2;
        int engExamDaysCol2 = 3;
        int engPCCountCol2 = 4;
        int marathiExamDaysCol2 = 6;
        int marathiPCCountCol2 = 7;
        int engMarathiExamDaysCol2 = 10;
        int totalDaysExamCol2 = 12;
        int Total_Amount2=17;
        // Result column (you can choose an empty column, e.g., 20th column)
        int resultCol2 = 20;
        int mismatchCol2 = 21;
        int WrongAmount2=22;
// Create headers
        Row headerRow2 = sheet2.getRow(0);
        if (headerRow2 == null) headerRow2 = sheet2.createRow(0);
        headerRow2.createCell(resultCol2).setCellValue("Result");
        headerRow2.createCell(mismatchCol2).setCellValue("Mismatch Details");
        headerRow2.createCell(WrongAmount2).setCellValue("Amount");

// Iterate rows
        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            Row row1 = sheet2.getRow(i);
            if (row1 == null) continue;

            String cCode1 = getValue(row1.getCell(centreCodeCol2));
            String cName1 = getValue(row1.getCell(centreNameCol2));
            String engDays1 = getValue(row1.getCell(engExamDaysCol2));
            String engPcCnt1 = getValue(row1.getCell(engPCCountCol2));
            String marDays1 = getValue(row1.getCell(marathiExamDaysCol2));
            String marPcCnt1 = getValue(row1.getCell(marathiPCCountCol2));
            String engMarDays1 = getValue(row1.getCell(engMarathiExamDaysCol2));
            String totalDays1 = getValue(row1.getCell(totalDaysExamCol2));
            String TotalAmountExl=getValue(row1.getCell(Total_Amount2));

            int ExlAmt2 = 0;
            int TtlDAys2 = 0;

            if (!TotalAmountExl.isEmpty()) {
                ExlAmt2 = Integer.parseInt(TotalAmountExl);
            }
            if (!totalDays1.isEmpty()) {
                TtlDAys2 = Integer.parseInt(totalDays1);
            }
            Cell amt = row1.getCell(WrongAmount2);
            if (amt == null) amt = row1.createCell(WrongAmount2); // Create if null

            if (ExlAmt2 != (TtlDAys2 * 600)) {
                amt.setCellValue(TtlDAys2 * 600);
            } else {
                amt.setCellValue(""); // Optional: clear if matches
            }
            boolean found = false;
            String result = "";
            String mismatchDetails = "";

            for (CollegeRecord rec : list) {
                if (cCode1.equals(rec.centreCode)) {
                    found = true;
                    List<String> mismatches = new ArrayList<>();

                    if (!cName1.equals(rec.centreName)) mismatches.add("CentreName");
                    if (!engDays1.equals(rec.englishExamDays)) mismatches.add("EnglishExamDays");
                    if (!engPcCnt1.equals(rec.englishPCCount)) mismatches.add("EnglishPCCount");
                    if (!marDays1.equals(rec.marathiExamDays)) mismatches.add("MarathiExamDays");
                    if (!marPcCnt1.equals(rec.marathiPCCount)) mismatches.add("MarathiPCCount");
                    if (!engMarDays1.equals(rec.englishMarathiExamDays)) mismatches.add("EnglishMarathiExamDays");
                    if (!totalDays1.equals(rec.totalDaysWithMockTest)) mismatches.add("TotalDaysWithMockTest");

                    if (mismatches.isEmpty()) {
                        result = "Match";
                    } else {
                        result = "Mismatch";
                        mismatchDetails = String.join(", ", mismatches);
                    }
                    break;
                }
            }

            if (!found) {
                result = "Not Found";
            }

            // Write to Excel
            Cell resultCell = row1.getCell(resultCol);
            if (resultCell == null) resultCell = row1.createCell(resultCol);
            resultCell.setCellValue(result);

            Cell mismatchCell = row1.getCell(mismatchCol);
            if (mismatchCell == null) mismatchCell = row1.createCell(mismatchCol);
            mismatchCell.setCellValue(mismatchDetails);
        }
        //test
        Sheet sheet3 = workbook.getSheet("Winner Technical");
        // Column indexes
        int centreCodeCol3 = 1;
        int centreNameCol3 = 2;
        int engExamDaysCol3 = 3;
        int engPCCountCol3 = 4;
        int marathiExamDaysCol3 = 6;
        int marathiPCCountCol3 = 7;
        int engMarathiExamDaysCol3 = 10;
        int totalDaysExamCol3 = 12;
        int Total_Amount3=17;
        // Result column (you can choose an empty column, e.g., 20th column)
        int resultCol3 = 20;
        int mismatchCol3 = 21;
        int WrongAmount3=22;
// Create headers
        Row headerRow3 = sheet3.getRow(0);
        if (headerRow3 == null) headerRow3 = sheet3.createRow(0);
        headerRow3.createCell(resultCol3).setCellValue("Result");
        headerRow3.createCell(mismatchCol3).setCellValue("Mismatch Details");
        headerRow3.createCell(WrongAmount3).setCellValue("Amount");

// Iterate rows
        for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
            Row row1 = sheet3.getRow(i);
            if (row1 == null) continue;

            String cCode1 = getValue(row1.getCell(centreCodeCol3));
            String cName1 = getValue(row1.getCell(centreNameCol3));
            String engDays1 = getValue(row1.getCell(engExamDaysCol3));
            String engPcCnt1 = getValue(row1.getCell(engPCCountCol3));
            String marDays1 = getValue(row1.getCell(marathiExamDaysCol3));
            String marPcCnt1 = getValue(row1.getCell(marathiPCCountCol3));
            String engMarDays1 = getValue(row1.getCell(engMarathiExamDaysCol3));
            String totalDays1 = getValue(row1.getCell(totalDaysExamCol3));
            String TotalAmountExl=getValue(row1.getCell(Total_Amount3));

            int ExlAmt2 = 0;
            int TtlDAys2 = 0;

            if (!TotalAmountExl.isEmpty()) {
                ExlAmt2 = Integer.parseInt(TotalAmountExl);
            }
            if (!totalDays1.isEmpty()) {
                TtlDAys2 = Integer.parseInt(totalDays1);
            }
            Cell amt = row1.getCell(WrongAmount3);
            if (amt == null) amt = row1.createCell(WrongAmount3); // Create if null

            if (ExlAmt2 != (TtlDAys2 * 400)) {
                amt.setCellValue(TtlDAys2 * 400);
            } else {
                amt.setCellValue(""); // Optional: clear if matches
            }
            boolean found = false;
            String result = "";
            String mismatchDetails = "";

            for (CollegeRecord rec : list) {
                if (cCode1.equals(rec.centreCode)) {
                    found = true;
                    List<String> mismatches = new ArrayList<>();

                    if (!cName1.equals(rec.centreName)) mismatches.add("CentreName");
                    if (!engDays1.equals(rec.englishExamDays)) mismatches.add("EnglishExamDays");
                    if (!engPcCnt1.equals(rec.englishPCCount)) mismatches.add("EnglishPCCount");
                    if (!marDays1.equals(rec.marathiExamDays)) mismatches.add("MarathiExamDays");
                    if (!marPcCnt1.equals(rec.marathiPCCount)) mismatches.add("MarathiPCCount");
                    if (!engMarDays1.equals(rec.englishMarathiExamDays)) mismatches.add("EnglishMarathiExamDays");
                    if (!totalDays1.equals(rec.totalDaysWithMockTest)) mismatches.add("TotalDaysWithMockTest");

                    if (mismatches.isEmpty()) {
                        result = "Match";
                    } else {
                        result = "Mismatch";
                        mismatchDetails = String.join(", ", mismatches);
                    }
                    break;
                }
            }

            if (!found) {
                result = "Not Found";
            }

            // Write to Excel
            Cell resultCell = row1.getCell(resultCol);
            if (resultCell == null) resultCell = row1.createCell(resultCol);
            resultCell.setCellValue(result);

            Cell mismatchCell = row1.getCell(mismatchCol);
            if (mismatchCell == null) mismatchCell = row1.createCell(mismatchCol);
            mismatchCell.setCellValue(mismatchDetails);
        }
        //test
        Sheet sheet4 = workbook.getSheet("Extra Winner Technical");
        // Column indexes
        int centreCodeCol4 = 1;
        int centreNameCol4 = 2;
        int engExamDaysCol4 = 3;
        int engPCCountCol4 = 4;
        int marathiExamDaysCol4 = 6;
        int marathiPCCountCol4 = 7;
        int engMarathiExamDaysCol4 = 10;
        int totalDaysExamCol4 = 12;
        int Total_Amount4=17;
        // Result column (you can choose an empty column, e.g., 20th column)
        int resultCol4 = 20;
        int mismatchCol4 = 21;
        int WrongAmount4=22;
// Create headers
        Row headerRow4 = sheet4.getRow(0);
        if (headerRow4 == null) headerRow4 = sheet4.createRow(0);
        headerRow4.createCell(resultCol4).setCellValue("Result");
        headerRow4.createCell(mismatchCol4).setCellValue("Mismatch Details");
        headerRow4.createCell(WrongAmount4).setCellValue("Amount");

// Iterate rows
        for (int i = 1; i <= sheet4.getLastRowNum(); i++) {
            Row row1 = sheet4.getRow(i);
            if (row1 == null) continue;

            String cCode1 = getValue(row1.getCell(centreCodeCol4));
            String cName1 = getValue(row1.getCell(centreNameCol4));
            String engDays1 = getValue(row1.getCell(engExamDaysCol4));
            String engPcCnt1 = getValue(row1.getCell(engPCCountCol4));
            String marDays1 = getValue(row1.getCell(marathiExamDaysCol4));
            String marPcCnt1 = getValue(row1.getCell(marathiPCCountCol4));
            String engMarDays1 = getValue(row1.getCell(engMarathiExamDaysCol4));
            String totalDays1 = getValue(row1.getCell(totalDaysExamCol4));
            String TotalAmountExl=getValue(row1.getCell(Total_Amount4));

            int ExlAmt2 = 0;
            int TtlDAys2 = 0;

            if (!TotalAmountExl.isEmpty()) {
                ExlAmt2 = Integer.parseInt(TotalAmountExl);
            }
            if (!totalDays1.isEmpty()) {
                TtlDAys2 = Integer.parseInt(totalDays1);
            }
            Cell amt = row1.getCell(WrongAmount3);
            if (amt == null) amt = row1.createCell(WrongAmount3); // Create if null

            if (ExlAmt2 != (TtlDAys2 * 400)) {
                amt.setCellValue(TtlDAys2 * 400);
            } else {
                amt.setCellValue(""); // Optional: clear if matches
            }
            boolean found = false;
            String result = "";
            String mismatchDetails = "";

            for (CollegeRecord rec : list) {
                if (cCode1.equals(rec.centreCode)) {
                    found = true;
                    List<String> mismatches = new ArrayList<>();

                    if (!cName1.equals(rec.centreName)) mismatches.add("CentreName");
                    if (!engDays1.equals(rec.englishExamDays)) mismatches.add("EnglishExamDays");
                    if (!engPcCnt1.equals(rec.englishPCCount)) mismatches.add("EnglishPCCount");
                    if (!marDays1.equals(rec.marathiExamDays)) mismatches.add("MarathiExamDays");
                    if (!marPcCnt1.equals(rec.marathiPCCount)) mismatches.add("MarathiPCCount");
                    if (!engMarDays1.equals(rec.englishMarathiExamDays)) mismatches.add("EnglishMarathiExamDays");
                    if (!totalDays1.equals(rec.totalDaysWithMockTest)) mismatches.add("TotalDaysWithMockTest");

                    if (mismatches.isEmpty()) {
                        result = "Match";
                    } else {
                        result = "Mismatch";
                        mismatchDetails = String.join(", ", mismatches);
                    }
                    break;
                }
            }

            if (!found) {
                result = "Not Found";
            }

            // Write to Excel
            Cell resultCell = row1.getCell(resultCol);
            if (resultCell == null) resultCell = row1.createCell(resultCol);
            resultCell.setCellValue(result);

            Cell mismatchCell = row1.getCell(mismatchCol);
            if (mismatchCell == null) mismatchCell = row1.createCell(mismatchCol);
            mismatchCell.setCellValue(mismatchDetails);
        }
        /// test
        fis.close();

        // Write updated workbook to file
        FileOutputStream fos = new FileOutputStream("F:\\GCC-TBC October  repeaters\\Remuneration_Updated_Result123.xlsx");
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Comparison complete. Results written to Excel.");
    }

    private static String getValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell))
                    return cell.getDateCellValue().toString();
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }
}


