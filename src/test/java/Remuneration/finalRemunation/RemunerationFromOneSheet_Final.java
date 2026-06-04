package Remuneration.finalRemunation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class RemunerationFromOneSheet_Final {
    static List<String> mismatchList = new ArrayList<>();
    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream(
                new File("F:\\GCCTBC-APR 2026\\Remunation\\Final\\All Region REMUNERATION_1.xlsx")
        );
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("Sheet1");
        int centreCodeCol = 1;
        //  HashMap to group rows by Centre Code
        Map<String, List<Row>> map = new HashMap<>();
        // STEP 1  READ ALL RECORDS & GROUP THEM
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String centreCode = getValue(row.getCell(centreCodeCol));
            if (centreCode.isEmpty()) continue;
           map.computeIfAbsent(centreCode, k -> new ArrayList<>()).add(row);
        }
        // STEP 2  PROCESS EACH CENTRE CODE
        for (String centreCode : map.keySet()) {

            List<Row> rows = map.get(centreCode);
            System.out.println("\n-------------------------------");
            System.out.println("CENTRE CODE: " + centreCode);
            System.out.println("TOTAL ROWS: " + rows.size());
            System.out.println("-------------------------------");
            //  validations per centre code group
            validateCentre(centreCode, rows);
        }

        workbook.close();
        fis.close();
        writeMismatchExcel();
        System.out.println("\n✔ Mismatch Excel Successfully Generated!");
    }
    // CONDITIONS
    static void validateCentre(String centreCode, List<Row> rows) {
        int centreNameCol = 2;
        int engDaysCol = 3;
        int engPcCol = 4;
        int engStdCol=5;
        int marDaysCol = 6;
        int marPcCol = 7;
        int marStdCol=8;
        int engMarDaysCol = 10;
        int totalDaysCol = 12;
        int designationCol=16;
        int totalAmtCol=19;
        int mockday=11;
        // Condition 1: All rows must have same centre name
        String firstName = getValue(rows.get(0).getCell(centreNameCol));

        for (Row r : rows) {
            String name = getValue(r.getCell(centreNameCol));
            if (!name.equals(firstName)) {
                System.out.println("❌ Name mismatch for centre " + centreCode);
                addMismatch(centreCode, "Name mismatch | Row " + r.getRowNum());
            }
        }
        // Condition 2: Sum all English PC Count
        int totalEngPC = 0;
        for (Row r : rows) {
            totalEngPC += parseInt(getValue(r.getCell(engPcCol)));
        }
        System.out.println("Total English PC Count = " + totalEngPC);
//Condition 3: for more that 30 count
         String engPCct=getValue(rows.get(0).getCell(engPcCol));
        String marPCct=getValue(rows.get(0).getCell(marPcCol));
      int  EngPCCount= parseInt(engPCct);
      int MarPCCount= parseInt(marPCct);
      int maxPCCount= Math.max(EngPCCount,MarPCCount);
      int requiredmanPower=(int) Math.ceil(maxPCCount/30.0);
        int requiredmancenterHead= (int) Math.ceil(maxPCCount/50.0);
      int itTeacherCount = 0;
        for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("IT TEACHER")) {
                itTeacherCount++;
            }
        }
        System.out.println("IT Teacher Rows = " + itTeacherCount);
        if (itTeacherCount > 0) {
            if ((requiredmanPower != itTeacherCount) && itTeacherCount>requiredmanPower) {
                System.out.println("❌ IT Teacher manpower mismatch for centre " + centreCode
                        + " | Required = " + requiredmanPower
                        + " | Provided = " + itTeacherCount);
                addMismatch(centreCode, "IT Teacher manpower mismatch for PC count " + maxPCCount + " | Required=" + requiredmanPower + " | Provided=" + itTeacherCount);
            } else {
                System.out.println("✔ IT Teacher manpower is correct.");
            }
        }
        for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("IT TEACHER")) {
                int totalDays1 = parseInt(getValue(r.getCell(totalDaysCol)));
                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays1 * 700;
                if (rowAmount != expected) {
                    System.out.println("❌ Incorrect IT Teacher amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount);
                    addMismatch(centreCode, " Incorrect IT Teacher amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount);
                } else {
                    System.out.println("✔ Correct IT Teacher amount at row " + r.getRowNum());
                }
            }
        }
        int winnerTechnical=0;
        for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("WINNER TECHNICAL")) {
                winnerTechnical++;
            }
        }
        System.out.println("Winner Technical Rows = " + winnerTechnical);
        if(winnerTechnical>0){
            if((requiredmanPower!=winnerTechnical)&&requiredmanPower<winnerTechnical)
            {
                System.out.println("❌ Winner Technical manpower mismatch for centre " + centreCode
                        + " | Required = " + requiredmanPower
                        + " | Provided = " + winnerTechnical);
                addMismatch(centreCode," Winner Technical manpower mismatch for centre " + centreCode
                        + " PC Count "+ maxPCCount + " | Required = " + requiredmanPower
                        + " | Provided = " + winnerTechnical);
            }
            else {
                System.out.println("✔ Winner Technical manpower is correct.");
            }
        }for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("WINNER TECHNICAL")) {
                int totalDays1 = parseInt(getValue(r.getCell(totalDaysCol)));
                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays1 * 450;
                if (rowAmount != expected) {
                    System.out.println("❌ Incorrect Winner Technical amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount);
                    addMismatch(centreCode," Incorrect Winner Technical amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount );
                } else {
                    System.out.println("✔ Correct Winner Technical amount at row " + r.getRowNum());
                }
            }
        }
        int peon=0;
        for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("PEON")) {
                peon++;
            }
        }
            if(peon>0){
                if((requiredmanPower!=peon)&&(peon>requiredmanPower))
                {
                    System.out.println("❌ Peon manpower mismatch for centre " + centreCode
                            + " | Required = " + requiredmanPower
                            + " | Provided = " + peon);
                    addMismatch(centreCode," Peon manpower mismatch for centre " + centreCode
                            +" PC Count "+ maxPCCount +  " | Required = " + requiredmanPower
                            + " | Provided = " + peon);
                }
                else {
                    System.out.println("✔ Peon manpower is correct.");
                }
            }
               for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("PEON")) {
                int engday=parseInt(getValue(r.getCell(engDaysCol)));
                int marday= parseInt(getValue(r.getCell(marDaysCol)));
               int mockD= parseInt(getValue(r.getCell(mockday)));
                int tot=parseInt(getValue(r.getCell(totalDaysCol)));
                int totalDays=tot-mockD;
                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays * 150;
                if (rowAmount != expected) {
                    System.out.println("❌ Incorrect Peon amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount);
                    addMismatch(centreCode," Incorrect Peon amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount );
                } else {
                    System.out.println("✔ Correct Peon amount at row " + r.getRowNum());
                }
            }
        }
int centerHead=0;
        for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("CENTER HEAD")) {
                centerHead++;
            }
        }
        if(centerHead>1)
        {
            if((requiredmancenterHead!=centerHead)&&centerHead>requiredmancenterHead)
            {
                System.out.println("❌ CENTER HEAD manpower mismatch for centre " + centreCode
                        + " | Required = " + requiredmancenterHead
                        + " | Provided = " + centerHead);
                addMismatch(centreCode, " CENTER HEAD manpower mismatch for centre " + centreCode
                        + " PC Count " + maxPCCount + " | Required = " + requiredmancenterHead
                        + " | Provided = " + centerHead);
            }
        } else if (centerHead == 0) {
            System.out.println("❌ CENTER HEAD manpower mismatch for centre " + centreCode
                    +" PC Count "+ maxPCCount +  " | Required = " + requiredmancenterHead
                    + " | Provided = " + centerHead);
            addMismatch(centreCode," CENTER HEAD manpower mismatch for centre " + centreCode
                    + " PC Count "+ maxPCCount + " | Required = " + requiredmancenterHead
                    + " | Provided = " + centerHead);
        } else {
            System.out.println("✔ CENTER HEAD manpower is correct.");
        }

        for (Row r : rows) {
            String desg = getValue(r.getCell(designationCol));
            if (desg.contains("CENTER HEAD")) {
                int engday=parseInt(getValue(r.getCell(engDaysCol)));
                int marday= parseInt(getValue(r.getCell(marDaysCol)));
                int mockD= parseInt(getValue(r.getCell(mockday)));
                int tot=parseInt(getValue(r.getCell(totalDaysCol)));
                int totalDays=tot-mockD;
                int rowAmount = parseInt(getValue(r.getCell(totalAmtCol)));
                int expected = totalDays * 700;
                if (rowAmount != expected) {
                    System.out.println("❌ Incorrect CENTER HEAD amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount);
                    addMismatch(centreCode," Incorrect CENTER HEAD amount at Centre " + centreCode
                            + " | Row#: " + r.getRowNum()
                            + " | Expected = " + expected
                            + " | Found = " + rowAmount );
                } else {
                    System.out.println("✔ Correct CENTER HEAD amount at row " + r.getRowNum());
                }
            }
        }
    }
    // Add mismatch to list
    static void addMismatch(String centreCode, String msg) {
        mismatchList.add(centreCode + " | " + msg);
    }

    // WRITE TO EXCEL
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
        FileOutputStream fos = new FileOutputStream("F:\\GCCTBC-APR 2026\\Remunation\\Final\\Mismatch_Report remuneration02-05-2026.xlsx");
        wb.write(fos);
        fos.close();
        wb.close();
    }
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
