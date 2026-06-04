package ExcelFileC6CellText;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadSpeedFirstThreeWordsFromDocToExcel {

    public static void main(String[] args) {
        // Input folder containing Word files
        File folder = new File("F:\\GCCTBC-APR 2026\\Mar_3040_AnswerFiles_Apr_2026");
        // Output Excel file
        String outputExcel = "F:\\GCCTBC-APR 2026\\Wrong compare check\\marathi\\Speed First 3 WordsOutput.xlsx";

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("First3Words");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("File Name");
            header.createCell(1).setCellValue("First 3 Words");

            int rowNum = 1;

            for (File file : folder.listFiles()) {
                if (file.isFile() && (file.getName().endsWith(".doc") || file.getName().endsWith(".docx"))&&(file.getName().startsWith("SpeedAnswer"))) {
                    String firstThreeWords = extractFirstThreeWords(file);
                    System.out.println(file);
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(file.getName());
                    row.createCell(1).setCellValue(firstThreeWords);
                }
            }

            // Write to Excel
            try (FileOutputStream fos = new FileOutputStream(outputExcel)) {
                workbook.write(fos);
            }

            System.out.println(" Excel created successfully at: " + outputExcel);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Extract only first 3 words from first non-empty paragraph
    private static String extractFirstThreeWords(File file) {
        try {
            String text = "";
            if ((file.getName().endsWith(".docx"))&&(file.getName().startsWith("SpeedAnswer"))) {
                try (FileInputStream fis = new FileInputStream(file);
                     XWPFDocument doc = new XWPFDocument(fis)) {
                    for (XWPFParagraph para : doc.getParagraphs()) {
                        text = para.getText().trim();
                        if (!text.isEmpty()) break;
                    }
                }
            } else if ((file.getName().endsWith(".doc"))&&(file.getName().startsWith("SpeedAnswer"))) {
                try (FileInputStream fis = new FileInputStream(file);
                     HWPFDocument doc = new HWPFDocument(fis);
                     WordExtractor extractor = new WordExtractor(doc)) {
                    for (String para : extractor.getParagraphText()) {
                        text = para.trim();
                        if (!text.isEmpty()) break;
                    }
                }
            }

            if (text.isEmpty()) return "(No text found)";

            // Get only first 3 words
            String[] words = text.split("\\s+");
            int limit = Math.min(words.length, 3);
            StringBuilder first3 = new StringBuilder();
            for (int i = 0; i < limit; i++) {
                first3.append(words[i]);
                if (i < limit - 1) first3.append(" ");
            }
            return first3.toString();

        } catch (Exception e) {
            System.err.println("Error reading file: " + file.getName() + " → " + e.getMessage());
        }
        return "(Error)";
    }
}
