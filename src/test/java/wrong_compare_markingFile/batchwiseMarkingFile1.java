package wrong_compare_markingFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashSet;

public class batchwiseMarkingFile1 {

    public static void main(String[] args) throws Exception {

        File batchwise = new File("F:\\GCC  TBC December 2025\\Question\\allocation\\marathi\\Allocation\\Batch Wise Subjective.xlsx");   // MASTER
        File studentbatchwise = new File("F:\\GCC  TBC December 2025\\Marking\\Marathi 30\\WrongComMarathihin30.xlsx"); // STUDENT

        DataFormatter df = new DataFormatter();

        HashSet<String> bset = new HashSet<>();

        FileInputStream fis = new FileInputStream(batchwise);
        Workbook wb = WorkbookFactory.create(fis);
        Sheet bsheet = wb.getSheetAt(0);

        for (int i = 1; i <= bsheet.getLastRowNum(); i++) {
            Row brow = bsheet.getRow(i);
            if (brow == null) continue;

            String batch = df.formatCellValue(brow.getCell(6)).trim();
            String filename = df.formatCellValue(brow.getCell(2)).trim();

            bset.add(batch + " || " + filename);
        }
        wb.close();

        // READ STUDENT
        FileInputStream fis1 = new FileInputStream(studentbatchwise);
        Workbook wb1 = WorkbookFactory.create(fis1);
        Sheet ssheet = wb1.getSheetAt(0);

        // OUTPUT FILE
        Workbook outWb = new XSSFWorkbook();
        Sheet matchSheet = outWb.createSheet("MATCHED");
        Sheet mismatchSheet = outWb.createSheet("MISMATCHED");

        createHeader(matchSheet);
        createHeader(mismatchSheet);

        int mRow = 1, mmRow = 1;

        for (int i = 1; i <= ssheet.getLastRowNum(); i++) {
            Row srow = ssheet.getRow(i);
            if (srow == null) continue;

            String seatNumber = df.formatCellValue(srow.getCell(0));
            System.out.println(seatNumber);

            // Letter
            checkAndWrite(seatNumber,
                    df.formatCellValue(srow.getCell(3)),
                    "LetterQuestion",
                    df.formatCellValue(srow.getCell(8)),
                    bset, matchSheet, mismatchSheet, mRow++, mmRow++);

            // Statement
            checkAndWrite(seatNumber,
                    df.formatCellValue(srow.getCell(5)),
                    "StatmentQuestion",
                    df.formatCellValue(srow.getCell(9)),
                    bset, matchSheet, mismatchSheet, mRow++, mmRow++);

            // Speed
            checkAndWrite(seatNumber,
                    df.formatCellValue(srow.getCell(7)),
                    "SpeedQuestion",
                    df.formatCellValue(srow.getCell(10)),
                    bset, matchSheet, mismatchSheet, mRow++, mmRow++);
        }

        wb1.close();

        FileOutputStream fos = new FileOutputStream("F:\\GCC  TBC December 2025\\Marking\\Marathi 30\\Final_Match_Report.xlsx");
        outWb.write(fos);
        fos.close();
        outWb.close();

        System.out.println(" Matching completed successfully");
    }

    //  MATCH CHECK
    static void checkAndWrite(String seat, String batch, String type, String file,
                              HashSet<String> bset,
                              Sheet matchSheet, Sheet mismatchSheet,
                              int mRow, int mmRow) {

        if (file == null || file.isEmpty()) return;

        boolean match = bset.contains(batch + " || " + file);

        Sheet target = match ? matchSheet : mismatchSheet;
        int rowNum = match ? mRow : mmRow;

        Row r = target.createRow(rowNum);
        r.createCell(0).setCellValue(seat);
        r.createCell(1).setCellValue(batch);
        r.createCell(2).setCellValue(type);
        r.createCell(3).setCellValue(file);
        r.createCell(4).setCellValue(match ? "MATCH" : "MISMATCH");
    }

    static void createHeader(Sheet s) {
        Row h = s.createRow(0);
        h.createCell(0).setCellValue("SeatNumber");
        h.createCell(1).setCellValue("BatchId");
        h.createCell(2).setCellValue("QuestionType");
        h.createCell(3).setCellValue("QuestionFile");
        h.createCell(4).setCellValue("Status");
    }
}
