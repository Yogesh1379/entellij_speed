package Check_Wrong_Comapre;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

public class Compare_First10Words_Sequence {

    public static void main(String[] args) throws IOException {

        // 🔹 Folder containing student files
        File studentFolder = new File("F:\\GCC-TBC October  repeaters\\Answer Files\\SpeedAnsFiles");
        File[] docxFiles = studentFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx"));

        if (docxFiles == null || docxFiles.length == 0) {
            System.out.println("No student files found!");
            return;
        }

        // 🔹 Excel allocation file (same logic from your previous codes)
        FileInputStream fisAlloc = new FileInputStream(
                "F:\\GCC-TBC October  repeaters\\eng allocation\\Batch Wise Subjective (1).xlsx");
        Workbook allocWorkbook = new XSSFWorkbook(fisAlloc);
        Sheet allocSheet = allocWorkbook.getSheetAt(0);

        // 🔹 Create result workbook
        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("First10WordMismatch");

        // Header row
        Row header = resultSheet.createRow(0);
        header.createCell(0).setCellValue("Seat No");
        header.createCell(1).setCellValue("Student File");
        header.createCell(2).setCellValue("Model File");
        header.createCell(3).setCellValue("Model 10 Words");
        header.createCell(4).setCellValue("Student 10 Words");
        header.createCell(5).setCellValue("Matched Words");

        int rowNum = 1;

        for (File studentFile : docxFiles) {
            String stdFileName = studentFile.getName();
            String[] parts = stdFileName.split("_");
            if (parts.length < 3) continue;

            String seatno = parts[1];
            String batchname = parts[2];

            // Determine course from seat number
            int course1 = 0;
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                if (course.equals("15")) course1 = 1;
                else if (course.equals("16")) course1 = 2;
            }

            // 🔹 Find model file using allocation Excel
            File modelFile = null;
            int rowCount = 0;
            for (Row row : allocSheet) {
                if (rowCount++ == 0) continue;
                Cell batchCell = row.getCell(6);
                Cell courseCell = row.getCell(5);
                Cell subjectiveCell = row.getCell(2);
                if (batchCell == null || courseCell == null || subjectiveCell == null) continue;

                String excelBatch = getCellValueAsString(batchCell);
                int excelCourse1 = Integer.parseInt(getCellValueAsString(courseCell));

                if (excelBatch.equals(batchname) && excelCourse1 == course1) {
                    String fileCandidate = getCellValueAsString(subjectiveCell);
                    if ((fileCandidate.startsWith("Eng30 Speed") && course1 == 1)
                            || (fileCandidate.startsWith("Eng 40 Speed") && course1 == 2)) {
                        modelFile = new File(
                                "F:\\GCC-TBC October  repeaters\\Question files\\Speed Answer files\\" + fileCandidate);
                        break;
                    }
                }
            }

            if (modelFile == null || !modelFile.exists()) {
                System.out.println("No model file for: " + stdFileName);
                continue;
            }

            // 🔹 Extract first 10 words from each
            String[] modelWords = extractFirstNWords(modelFile, 5);
            String[] studentWords = extractFirstNWords(studentFile, 5);

            String modelStr = String.join(" ", modelWords);
            String studentStr = String.join(" ", studentWords);

            // 🔹 Count sequence matches
            int matchedCount = countSequentialMatches(modelWords, studentWords);

            // ✅ Require at least 5 in correct sequence
            if (matchedCount < 5) {
                Row r = resultSheet.createRow(rowNum++);
                r.createCell(0).setCellValue(seatno);
                r.createCell(1).setCellValue(stdFileName);
                r.createCell(2).setCellValue(modelFile.getName());
                r.createCell(3).setCellValue(modelStr);
                r.createCell(4).setCellValue(studentStr);
                r.createCell(5).setCellValue(matchedCount);
                System.out.println("❌ Mismatch (" + matchedCount + " matches): " + stdFileName);
            }
        }

        // 🔹 Save result Excel
        FileOutputStream fos = new FileOutputStream(
                new File(studentFolder.getParentFile(), "First10WordMismatchSummary.xlsx"));
        resultWorkbook.write(fos);
        fos.close();
        resultWorkbook.close();
        allocWorkbook.close();
        fisAlloc.close();

        System.out.println("\n✅ Comparison completed. Result saved in First10WordMismatchSummary.xlsx");
    }

    // Extracts the first N words from the DOCX file (ignores blank lines)
    private static String[] extractFirstNWords(File file, int n) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis);
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph p : doc.getParagraphs()) {
            String text = p.getText().trim();
            if (!text.isEmpty()) {
                sb.append(text).append(" ");
                if (sb.toString().split("\\s+").length >= n) break;
            }
        }
        fis.close();

        String[] words = sb.toString().trim().replaceAll("\\s+", " ").split(" ");
        if (words.length > n) words = Arrays.copyOfRange(words, 0, n);
        return Arrays.stream(words)
                .map(Compare_First10Words_Sequence::normalize)
                .toArray(String[]::new);
    }

    // Count how many words match in correct sequence
    private static int countSequentialMatches(String[] modelWords, String[] studentWords) {
        int matched = 0;
        int j = 0;
        for (String mw : modelWords) {
            while (j < studentWords.length) {
                if (mw.equalsIgnoreCase(studentWords[j])) {
                    matched++;
                    j++;
                    break;
                }
                j++;
            }
        }
        return matched;
    }

    private static String getCellValueAsString(Cell cell) {
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

    static String normalize(String s) {
        if (s == null) return "";
        return s.replaceAll("\\s+", " ").trim()
                .replace("’", "'")
                .replace("‘", "'")
                .replace("“", "\"")
                .replace("”", "\"");
    }
}
