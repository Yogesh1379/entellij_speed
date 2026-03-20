package Final_SpeedMarking;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FinalMarathi_EnterAndTabCount_1 {

    static int FormatingMistake = 0;
    static int spaceCount = 0;
    static int tabCount = 0;
    static int maxBlankSequence = 0;

    public static void main(String[] args) throws IOException {

        // Folder containing student files
        File studentFolder = new File("F:\\GCC-TBC October  repeaters\\Answer Files\\SpeedAnsFiles");
        File[] docxFiles = studentFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx"));

        if (docxFiles == null || docxFiles.length == 0) {
            System.out.println("No student files found!");
            return;
        }

        // Output folder for highlighted DOCX
        File outputFolder = new File(studentFolder.getParentFile(), "Results_Highlighted");
        if (!outputFolder.exists()) outputFolder.mkdirs();

        // Create Excel summary workbook
        Workbook summaryWorkbook = new XSSFWorkbook();
        Sheet summarySheet = summaryWorkbook.createSheet("Summary");

        // Header row
        Row header = summarySheet.createRow(0);
        header.createCell(0).setCellValue("Seat No");
        header.createCell(1).setCellValue("Student File");
        header.createCell(2).setCellValue("Total Mistakes");
        header.createCell(3).setCellValue("Obtained Marks");
        header.createCell(4).setCellValue("Extra Word");
        header.createCell(5).setCellValue("Missing Word");
        header.createCell(6).setCellValue("Wrong Word");
        header.createCell(7).setCellValue("Extra Enter");
        header.createCell(8).setCellValue("Extra Tab");
        header.createCell(9).setCellValue("Extra Space");
        header.createCell(10).setCellValue("All Mistakes");

        int rowNum = 1;

        // Load allocation Excel only once
        FileInputStream fisAlloc = new FileInputStream(
                "F:\\GCC-TBC October  repeaters\\Marathi\\allocation\\Batch Wise Subjective (1).xlsx");
        Workbook allocWorkbook = new XSSFWorkbook(fisAlloc);
        Sheet allocSheet = allocWorkbook.getSheetAt(0);

        for (File studentFile : docxFiles) {
            String stdfile = studentFile.getName();
            String[] parts = stdfile.split("_");
            if (parts.length < 3) continue;

            String seatno = parts[1];
            String batchname = parts[2];
            int course1 = 0;
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                if (course.equals("25")) course1 = 3;
                else if (course.equals("26")) course1 = 4;
                else if (course.equals("35")) course1 = 5;
                else if (course.equals("36")) course1 = 6;
            }

            // Determine model file
            File modelFile = null;
            int rowCount = 0;
            for (Row row : allocSheet) {
                if (rowCount++ == 0) continue; // skip header
                Cell batchCell = row.getCell(6);
                Cell courseCell = row.getCell(5);
                Cell subjectiveCell = row.getCell(2);
                if (batchCell == null || courseCell == null || subjectiveCell == null) continue;

                String excelBatch = getCellValueAsString(batchCell);
                int excelCourse1 = Integer.parseInt(getCellValueAsString(courseCell));

                if (excelBatch.equals(batchname) && excelCourse1 == course1) {
                    String fileCandidate = getCellValueAsString(subjectiveCell);
                    if ((fileCandidate.startsWith("Mar30 Speed") && course1 == 3) ||
                            (fileCandidate.startsWith("Mar 40 Speed") && course1 == 4)||
                    (fileCandidate.startsWith("Hin30 Speed") && course1 == 5)||
                    (fileCandidate.startsWith("Hin 40 Speed") && course1 == 6)){
                        modelFile = new File(
                                "F:\\GCC-TBC October  repeaters\\Question files\\Marathi  Speed model answer\\" + fileCandidate);
                        break;
                    }
                }
            }

            if (modelFile == null || !modelFile.exists()) {
                System.out.println("No model file for: " + stdfile);
                continue;
            }

            // Extract text from model and student files
            String modelText = extractFullText(modelFile);
            String studentText = extractFullText(studentFile);

            // Compare word by word and generate highlighted DOCX
            XWPFDocument outDoc = new XWPFDocument();
            XWPFParagraph para = outDoc.createParagraph();

            String[] mWords = modelText.split("\\s+");
            String[] sWords = studentText.split("\\s+");

            int i = 0, j = 0;
            int missingWordCount = 0, extraWordCount = 0, wrongWordCount = 0;
            StringBuilder allMistakes = new StringBuilder();

            while (i < mWords.length || j < sWords.length) {
                String refWord = i < mWords.length ? mWords[i] : null;
                String stuWord = j < sWords.length ? sWords[j] : null;

//                if (refWord != null && stuWord != null && refWord.equals(stuWord)) {
                if (refWord != null && stuWord != null && normalizeWord(refWord).equals(normalizeWord(stuWord))) {
                    XWPFRun r = para.createRun();
                    r.setText(stuWord + " ");
                    i++; j++;
                } else if (refWord != null && stuWord != null) {
                    if (j + 1 < sWords.length && normalizeWord(refWord).equals(normalizeWord(sWords[j + 1]))){
                        XWPFRun r = para.createRun();
                        r.setText("[ " + stuWord + "] ");
                        r.setColor("0000FF");
                        extraWordCount++;
                        allMistakes.append("Extra: ").append(stuWord).append("; ");
                        j++;
                    } else if (i + 1 < mWords.length && normalizeWord(mWords[i + 1]).equals(normalizeWord(stuWord))){
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWord + " ] ");
                        r.setColor("008000");
                        missingWordCount++;
                        allMistakes.append("Missing: ").append(refWord).append("; ");
                        i++;
                    } else {
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWord + " / " + stuWord + "] ");
                        r.setColor("FF0000");
                        wrongWordCount++;
                        allMistakes.append("Wrong: ").append(stuWord).append("; ");
                        i++; j++;
                    }
                } else {
                    if (refWord != null) {
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWord + " ] ");
                        r.setColor("008000");
                        missingWordCount++;
                        allMistakes.append("Missing: ").append(refWord).append("; ");
                        i++;
                    }
                    if (stuWord != null) {
                        XWPFRun r = para.createRun();
                        r.setText("[ " + stuWord + "] ");
                        r.setColor("0000FF");
                        extraWordCount++;
                        allMistakes.append("Extra: ").append(stuWord).append("; ");
                        j++;
                    }
                }
            }

            // Total formatting mistakes
            FormatingMistake = maxBlankSequence + tabCount + spaceCount;
            int totalMistakes = missingWordCount + extraWordCount + wrongWordCount + FormatingMistake;
            int obtainedMarks = Math.max(40 - totalMistakes, 0);

            // Add summary in DOCX
            XWPFParagraph summaryPara = outDoc.createParagraph();
            XWPFRun summaryRun = summaryPara.createRun();
            summaryRun.setText("\n--- Mistake Summary ---\n");
            summaryRun.addCarriageReturn();
            summaryRun.setText("Missing Words (Green) : " + missingWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Extra Words (Blue) : " + extraWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Wrong Words (Red) : " + wrongWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Extra Enter (2+): " + maxBlankSequence);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Extra Tab (3+): " + tabCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Extra Space (3+): " + spaceCount);
            summaryRun.addCarriageReturn();
            summaryRun.addCarriageReturn();
            summaryRun.setBold(true);
            summaryRun.setText("Total Marks: " + obtainedMarks);

            // Save highlighted DOCX
            FileOutputStream fos = new FileOutputStream(new File(outputFolder, "Result_" + stdfile));
            outDoc.write(fos);
            fos.close();
            outDoc.close();

            // Write Excel summary
            Row row = summarySheet.createRow(rowNum++);
            row.createCell(0).setCellValue(seatno); // seat number from filename
            row.createCell(1).setCellValue(stdfile);
            row.createCell(2).setCellValue(totalMistakes);
            row.createCell(3).setCellValue(obtainedMarks);
            row.createCell(4).setCellValue(extraWordCount);
            row.createCell(5).setCellValue(missingWordCount);
            row.createCell(6).setCellValue(wrongWordCount);
            row.createCell(7).setCellValue(maxBlankSequence);
            row.createCell(8).setCellValue(tabCount);
            row.createCell(9).setCellValue(spaceCount);
            row.createCell(10).setCellValue(allMistakes.toString());

            System.out.println(stdfile);
        }

        // Save Excel summary
        FileOutputStream fosExcel = new FileOutputStream(
                new File(studentFolder.getParentFile(), "Marathi_StudentMistakesSummary.xlsx"));
        summaryWorkbook.write(fosExcel);
        fosExcel.close();
        summaryWorkbook.close();
        allocWorkbook.close();
        fisAlloc.close();

        System.out.println("✅ All files processed. Summary Excel created!");
    }

    // ✅ UPDATED LOGIC — ignores leading spaces/tabs before paragraph start
    private static String extractFullText(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis);
        StringBuilder sb = new StringBuilder();

        int totalTabCount = 0;
        int totalSpaceCount = 0;

        int continuousBlankParas = 0;
        maxBlankSequence = 0;

        for (XWPFParagraph p : doc.getParagraphs()) {
            String text = p.getText();

            // Count blank paragraphs (for Extra Enter)
            if (text.trim().isEmpty()) {
                continuousBlankParas++;
                if (continuousBlankParas >= 1) maxBlankSequence++;
            } else {
                continuousBlankParas = 0;
            }

            // 🔹 Ignore leading tabs/spaces before first visible character
            String textAfterLeading = text.replaceAll("^[\\t ]+", "");

            // Count extra spaces/tabs only AFTER text begins
            totalTabCount += countMatches(textAfterLeading, "\t{1,}");
//            totalSpaceCount += countMatches(textAfterLeading, " {3,}");
            // Count extra spaces beyond 2 continuous
            Matcher spaceMatcher = Pattern.compile(" {2,}").matcher(textAfterLeading);
            while (spaceMatcher.find()) {
                int len = spaceMatcher.end() - spaceMatcher.start();
                totalSpaceCount += (len - 2); // only spaces beyond 2 are extra
            }

            sb.append(text).append("\n ");
        }

        fis.close();

        // Save counts globally
        tabCount = totalTabCount;
        spaceCount = totalSpaceCount;

        // Return clean text for comparison
        return sb.toString().replaceAll("\\r?\\n", " ").replaceAll("\\s+", " ").trim();
    }

    private static int countMatches(String text, String regex) {
        Matcher m = Pattern.compile(regex).matcher(text);
        int count = 0;
        while (m.find()) count++;
        return count;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
    private static String normalizeWord(String word) {
        if (word == null) return null;

        // Replace curly and accented quotes with straight ones
        return word
                .replace("’", "'")   // curly apostrophe
                .replace("‘", "'")   // opening curly single quote
                .replace("´", "'")   // acute accent often used as apostrophe
                .replace("`", "'")   // grave accent
                .replace("“", "\"")  // opening curly double quote
                .replace("”", "\""); // closing curly double quote
    }
}
