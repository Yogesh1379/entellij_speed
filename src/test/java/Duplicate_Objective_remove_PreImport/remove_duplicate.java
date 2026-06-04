package Duplicate_Objective_remove_PreImport;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.testng.annotations.Test;

import java.io.*;
import java.util.HashSet;

public class remove_duplicate {
    @Test
    public void duplicate() throws IOException, InvalidFormatException {
        File inputfile = new File("C:\\Users\\User\\Desktop\\New folder\\Repeater Question files April 2026\\ENGLISH\\OBJ_30&40-ENGLISH\\30.xlsx");
        File outputFile = new File("C:\\Users\\User\\Desktop\\New folder\\Repeater Question files April 2026\\ENGLISH\\OBJ_30&40-ENGLISH\\uniqobj30.xlsx");
        FileInputStream fis = new FileInputStream(inputfile);
        XSSFWorkbook wb = XSSFWorkbookFactory.createWorkbook(OPCPackage.open(fis));
        Sheet sheet = wb.getSheetAt(0);
        HashSet<String> uniqueset = new HashSet<>();
        XSSFWorkbook whOut = new XSSFWorkbook();
        XSSFSheet sheetOut = whOut.createSheet("Clean Data");
        XSSFRow header = sheetOut.createRow(0);
        header.createCell(0).setCellValue("QUESTION_TEXT_ENG");
        header.createCell(1).setCellValue("QUESTION_TEXT_MAR");
        header.createCell(2).setCellValue("QUESTION_TEXT_HIN");
        header.createCell(3).setCellValue("ANSWER");
        header.createCell(4).setCellValue("SUBJECT_ID");
        header.createCell(5).setCellValue("ANSWER_OPTION_TEXT_ENGLISH_A");
        header.createCell(6).setCellValue("ANSWER_OPTION_TEXT_ENGLISH_b");
        header.createCell(7).setCellValue("ANSWER_OPTION_TEXT_ENGLISH_C");
        header.createCell(8).setCellValue("ANSWER_OPTION_TEXT_ENGLISH_D");
        header.createCell(9).setCellValue("ANSWER_OPTION_TEXT_MARATHI_A");
        header.createCell(10).setCellValue("ANSWER_OPTION_TEXT_MARATHI_B");
        header.createCell(11).setCellValue("ANSWER_OPTION_TEXT_MARATHI_C");
        header.createCell(12).setCellValue("ANSWER_OPTION_TEXT_MARATHI_D");
        header.createCell(13).setCellValue("ANSWER_OPTION_TEXT_HINDI_A");
        header.createCell(14).setCellValue("ANSWER_OPTION_TEXT_HINDI_B");
        header.createCell(15).setCellValue("ANSWER_OPTION_TEXT_HINDI_C");
        header.createCell(16).setCellValue("ANSWER_OPTION_TEXT_HINDI_D");
        DataFormatter df = new DataFormatter();
        int r = 1;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if(row==null)continue;
            String QUESTION_TEXT_ENG = df.formatCellValue(row.getCell(0));
            String QUESTION_TEXT_MAR = df.formatCellValue(row.getCell(1));
            String QUESTION_TEXT_HIN = df.formatCellValue(row.getCell(2));
            String ANSWER = df.formatCellValue(row.getCell(3));
            String SUBJECT_ID = df.formatCellValue(row.getCell(4));
            String ANSWER_OPTION_TEXT_ENGLISH_A = df.formatCellValue(row.getCell(5));
            String ANSWER_OPTION_TEXT_ENGLISH_b = df.formatCellValue(row.getCell(6));
            String ANSWER_OPTION_TEXT_ENGLISH_C = df.formatCellValue(row.getCell(7));
            String ANSWER_OPTION_TEXT_ENGLISH_D = df.formatCellValue(row.getCell(8));
            String ANSWER_OPTION_TEXT_MARATHI_A = df.formatCellValue(row.getCell(9));
            String ANSWER_OPTION_TEXT_MARATHI_B = df.formatCellValue(row.getCell(10));
            String ANSWER_OPTION_TEXT_MARATHI_C = df.formatCellValue(row.getCell(11));
            String ANSWER_OPTION_TEXT_MARATHI_D = df.formatCellValue(row.getCell(12));
            String ANSWER_OPTION_TEXT_HINDI_A = df.formatCellValue(row.getCell(13));
            String ANSWER_OPTION_TEXT_HINDI_B = df.formatCellValue(row.getCell(14));
            String ANSWER_OPTION_TEXT_HINDI_C = df.formatCellValue(row.getCell(15));
            String ANSWER_OPTION_TEXT_HINDI_D = df.formatCellValue(row.getCell(16));

            System.out.println(QUESTION_TEXT_ENG);

            String normalizeQUESTION_TEXT_ENG = normalize(QUESTION_TEXT_ENG);
            String normalizeANSWER_OPTION_TEXT_ENGLISH_A = normalize(ANSWER_OPTION_TEXT_ENGLISH_A);
            String normalizeANSWER_OPTION_TEXT_ENGLISH_b = normalize(ANSWER_OPTION_TEXT_ENGLISH_b);
            String normalizeANSWER_OPTION_TEXT_ENGLISH_C = normalize(ANSWER_OPTION_TEXT_ENGLISH_C);
            String normalizeANSWER_OPTION_TEXT_ENGLISH_D = normalize(ANSWER_OPTION_TEXT_ENGLISH_D);

            String key = normalizeQUESTION_TEXT_ENG + "|" + normalizeANSWER_OPTION_TEXT_ENGLISH_A + "|" +
                    normalizeANSWER_OPTION_TEXT_ENGLISH_b + "|" + normalizeANSWER_OPTION_TEXT_ENGLISH_C + "|" +
                    normalizeANSWER_OPTION_TEXT_ENGLISH_D;
            if (!uniqueset.contains(key)) {
                uniqueset.add(key);
                XSSFRow rowOut = sheetOut.createRow(r++);
                rowOut.createCell(0).setCellValue(QUESTION_TEXT_ENG);
                rowOut.createCell(1).setCellValue(QUESTION_TEXT_MAR);
                rowOut.createCell(2).setCellValue(QUESTION_TEXT_HIN);
                rowOut.createCell(3).setCellValue(ANSWER);
                rowOut.createCell(4).setCellValue(SUBJECT_ID);
                rowOut.createCell(5).setCellValue(ANSWER_OPTION_TEXT_ENGLISH_A);
                rowOut.createCell(6).setCellValue(ANSWER_OPTION_TEXT_ENGLISH_b);
                rowOut.createCell(7).setCellValue(ANSWER_OPTION_TEXT_ENGLISH_C);
                rowOut.createCell(8).setCellValue(ANSWER_OPTION_TEXT_ENGLISH_D);
                rowOut.createCell(9).setCellValue(ANSWER_OPTION_TEXT_MARATHI_A);
                rowOut.createCell(10).setCellValue(ANSWER_OPTION_TEXT_MARATHI_B);
                rowOut.createCell(11).setCellValue(ANSWER_OPTION_TEXT_MARATHI_C);
                rowOut.createCell(12).setCellValue(ANSWER_OPTION_TEXT_MARATHI_D);
                rowOut.createCell(13).setCellValue(ANSWER_OPTION_TEXT_HINDI_A);
                rowOut.createCell(14).setCellValue(ANSWER_OPTION_TEXT_HINDI_B);
                rowOut.createCell(15).setCellValue(ANSWER_OPTION_TEXT_HINDI_C);
                rowOut.createCell(16).setCellValue(ANSWER_OPTION_TEXT_HINDI_D);

            }


        }
        FileOutputStream fos = new FileOutputStream(outputFile);
        whOut.write(fos);
        fos.close();
        whOut.close();

        System.out.println("generated successfully!");


    }

    public static String normalize(String text) {
        if (text == null) return "";
        return text.toLowerCase()
                .replaceAll("[-_.]+", "")   // remove -, _, .
                .replaceAll("\\s+", " ")     // remove extra spaces
                .trim();
    }
}
