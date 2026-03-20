package Marathi_Ignore_Word;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

public class IgnoreWordEngineWithExcel {

    static Workbook ignoreWorkbook;
    static Workbook resultWorkbook;
    static Sheet resultSheet;
    static int resultRowNum = 1;

    public static void main(String[] args) throws Exception {

           //  STUDENT FOLDER PATH

        File studentFolder = new File(
                "F:\\GCC  TBC December 2025\\Mararthi bucket\\std ans file\\speed"
        );

        //  IGNORE EXCEL PATH
        String ignoreExcelPath =
                "F:\\GCC  TBC December 2025\\Marathi Ignore Word\\GCC TBC Mar30_40 Hindi 30_40 Speed ignore word_Final - Copy.xlsx";

        //  OUTPUT RESULT FILE
        String outputPath =
                "C:\\Users\\User\\Desktop\\New folder\\IgnoreWord_Result.xlsx";

        // Load Ignore Excel ONCE (Important for speed)
        FileInputStream ignoreFis = new FileInputStream(ignoreExcelPath);
        ignoreWorkbook = new XSSFWorkbook(ignoreFis);

        // Create Result Workbook
        resultWorkbook = new XSSFWorkbook();
        resultSheet = resultWorkbook.createSheet("Ignore Result");

        // Create Header
        Row header = resultSheet.createRow(0);
        header.createCell(0).setCellValue("Seat No");
        header.createCell(1).setCellValue("File Name");
        header.createCell(2).setCellValue("Batch");
        header.createCell(3).setCellValue("Matched Word");
        header.createCell(4).setCellValue("Extra Mark Count");

        File[] files = studentFolder.listFiles((dir, name) -> name.endsWith(".docx"));

        if (files == null) {
            System.out.println("No student files found.");
            return;
        }

        for (File file : files) {
            System.out.println("Processing file: " + file.getName());

            String seatNo = file.getName().replace(".docx", "");
            String fileName = file.getName();
//            String batchName = "Batch1";   // change if needed
//            int course1 = 3;               // change as per logic
//            String stdfile = studentFile.getName();
            String[] parts = fileName.split("_");
            if (parts.length < 3) continue;

            String seatno = parts[1];
            String batchName = parts[2];
            int course1 = 0;
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                course1 = switch (course) {
                    case "25" -> 3;
                    case "26" -> 4;
                    case "35" -> 5;
                    case "36" -> 6;
                    default -> course1;
                };
            }

            String studentText = extractDocText(file);

            String sheetName = getSheetName(course1);

            List<String> matchedIgnoreWords =
                    checkIgnoreWords(sheetName, batchName, studentText);

            int extraIgnoreMarkCount = matchedIgnoreWords.size();

            StringBuilder sb = new StringBuilder();

            Row  row = resultSheet.createRow(resultRowNum++);
            for (String word : matchedIgnoreWords) {


                sb.append(word).append(", ");
            }

            row.createCell(0).setCellValue(seatno);
            row.createCell(1).setCellValue(fileName);
            row.createCell(2).setCellValue(batchName);
            row.createCell(3).setCellValue(sb.toString());
            row.createCell(4).setCellValue(extraIgnoreMarkCount);

        }

        // Auto-size columns
        for (int i = 0; i <= 4; i++) {
            resultSheet.autoSizeColumn(i);
        }

        // Save result file
        FileOutputStream fos = new FileOutputStream(outputPath);
        resultWorkbook.write(fos);
        fos.close();

        ignoreWorkbook.close();
        resultWorkbook.close();

        System.out.println("Excel Created Successfully ✅");
    }

    // ===============================
    // Extract DOCX Text
    // ===============================
    private static String extractDocText(File file) throws Exception {

        StringBuilder text = new StringBuilder();

        FileInputStream fis = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fis);

        for (XWPFParagraph para : document.getParagraphs()) {
            text.append(para.getText()).append(" ");
        }

        document.close();
        fis.close();

        return text.toString();
    }

    // ===============================
    // Ignore Word Logic
    // ===============================
    private static List<String> checkIgnoreWords(
            String sheetName,
            String batchName,
            String studentText
    ) {

        List<String> matchedWords = new ArrayList<>();

        Sheet sheet = ignoreWorkbook.getSheet(sheetName);
        if (sheet == null) return matchedWords;

        // FAST SEARCH
        Set<String> studentWordSet = new HashSet<>();
        for (String w : studentText.split("\\s+")) {
            studentWordSet.add(normalize(w));
        }

        for (Row row : sheet) {

            if (row.getRowNum() == 0) continue;

            String batch = getCellValue(row.getCell(1));
            if (!batch.equalsIgnoreCase(batchName)) continue;

            String paperWord = normalize(getCellValue(row.getCell(2)));
            String studentWord = normalize(getCellValue(row.getCell(3)));

            if (!paperWord.isEmpty() && studentWordSet.contains(paperWord))
                continue;

            if (!studentWord.isEmpty() && studentWordSet.contains(studentWord)) {
                matchedWords.add(studentWord);
            }
        }

        return matchedWords;
    }

    // ===============================
    private static String getSheetName(int course1) {

        switch (course1) {
            case 3: return "Marathi 30";
            case 4: return "Marathi 40";
            case 5: return "Hindi 30";
            case 6: return "Hindi 40";
            default: return "";
        }
    }

    // ===============================
    private static String getCellValue(Cell cell) {

        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }

    // ===============================
    private static String normalize(String word) {

        if (word == null) return "";

        return word.replace("’", "'")
                .replace("‘", "'")
                .replace("“", "\"")
                .replace("”", "\"")
                .trim();
    }
}