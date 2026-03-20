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

public class AllStudentFinalCurrectWithExcel {

    public static void main(String[] args) throws IOException {

        // Folder containing student files
        File studentFolder = new File("C:\\Users\\User\\Desktop\\New folder (3)");
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
        header.createCell(0).setCellValue("Student File");
        header.createCell(1).setCellValue("Total Mistakes");
        header.createCell(2).setCellValue("Obtained Marks");
        header.createCell(3).setCellValue("Extra Word");
        header.createCell(4).setCellValue("Missing Word");
        header.createCell(5).setCellValue("Wrong Word");
        header.createCell(6).setCellValue("All Mistakes");

        int rowNum = 1;

        // Loop over all student files
        for (File studentFile : docxFiles) {
            String stdfile = studentFile.getName();
            String[] parts = stdfile.split("_");
            if (parts.length < 3) continue; // ensure filename has enough parts

            String seatno = parts[1];
            String batchname = parts[2];

            int course1 = 0;
            if (seatno.length() >= 10) {
                String course = seatno.substring(4, 6);
                if (course.equals("15")) {
                    course1 = 1;
                } else if (course.equals("16")) {
                    course1 = 2;
                }
            }

            // Determine model file from Excel allocation
            File modelFile = null;
            FileInputStream fis = new FileInputStream(
                    "F:\\GCC-TBC October  repeaters\\eng allocation\\Batch Wise Subjective (1).xlsx");
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = 0;

            for (Row row : sheet) {
                if (rowCount++ == 0) continue; // skip header
                Cell batchCell = row.getCell(6);
                Cell courseCell = row.getCell(5);
                Cell subjectiveCell = row.getCell(2);
                if (batchCell == null || courseCell == null || subjectiveCell == null) continue;

                String excelBatch = getCellValueAsString(batchCell);
                String excelCourse = getCellValueAsString(courseCell);

                int excelCourse1 = Integer.parseInt(excelCourse);

                if (excelBatch.equals(batchname) && excelCourse1 == course1) {
                    String fileCandidate = getCellValueAsString(subjectiveCell);
                    if ((fileCandidate.startsWith("Eng30 Speed") && course1 == 1) ||
                            (fileCandidate.startsWith("Eng 40 Speed") && course1 == 2)) {
                        modelFile = new File(
                                "F:\\GCC-TBC October  repeaters\\Question files\\Speed Answer files\\" + fileCandidate);
                        break;
                    }
                }
            }
            workbook.close();
            fis.close();

            if (modelFile == null || !modelFile.exists()) {
                System.out.println("No model file found for student: " + stdfile);
                continue;
            }

            // Extract text from model and student files
            String modelText = extractFullText(modelFile);
            String studentText = extractFullText(studentFile);

            // Compare word by word and generate highlighted result
            XWPFDocument outDoc = new XWPFDocument();
            XWPFParagraph para = outDoc.createParagraph();

            String[] mWords = modelText.split("\\s+");
            String[] sWords = studentText.split("\\s+");

            int i = 0, j = 0;
            int missingWordCount = 0, extraWordCount = 0, wrongWordCount = 0;
            int totalMarks = 0;

            StringBuilder allMistakes = new StringBuilder(); // Collect all mistakes

            while (i < mWords.length || j < sWords.length) {
                String refWord = i < mWords.length ? mWords[i] : null;
                String stuWord = j < sWords.length ? sWords[j] : null;

                if (refWord != null && stuWord != null && refWord.equals(stuWord)) {
                    XWPFRun r = para.createRun();
                    r.setText(stuWord + " ");
                    i++;
                    j++;
                } else if (refWord != null && stuWord != null) {
                    if (j + 1 < sWords.length && refWord.equals(sWords[j + 1])) {
                        XWPFRun r = para.createRun();
                        r.setText("[ " + stuWord + "] ");
                        r.setColor("0000FF"); // extra
                        extraWordCount++;
                        allMistakes.append("Extra: ").append(stuWord).append("; ");
                        j++;
                    } else if (i + 1 < mWords.length && mWords[i + 1].equals(stuWord)) {
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWord + " ] ");
                        r.setColor("008000"); // missing
                        missingWordCount++;
                        allMistakes.append("Missing: ").append(refWord).append("; ");
                        i++;
                    } else {
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWord + " / " + stuWord + "] ");
                        r.setColor("FF0000"); // wrong
                        wrongWordCount++;
                        allMistakes.append("Wrong: ").append(stuWord).append("; ");
                        i++;
                        j++;
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

            int mistakeCount = missingWordCount + extraWordCount + wrongWordCount;
            totalMarks = Math.max(40 - mistakeCount, 0);

            // Add summary to DOCX
            XWPFParagraph summaryPara = outDoc.createParagraph();
            summaryPara.setSpacingBefore(200);
            XWPFRun summaryRun = summaryPara.createRun();
            summaryRun.setText("\n--- Mistake Summary ---\n");
            summaryRun.addCarriageReturn();
            summaryRun.setText("Missing Words (Green) : " + missingWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Extra Words (Blue) : " + extraWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Wrong Words (Red) : " + wrongWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.addCarriageReturn();
            summaryRun.setBold(true);
            summaryRun.setText("Total Marks : " + totalMarks);

            // Save highlighted DOCX
            FileOutputStream fos = new FileOutputStream(new File(outputFolder, "Result_" + stdfile));
            outDoc.write(fos);
            fos.close();
            outDoc.close();

            // Write Excel summary row
            Row row = summarySheet.createRow(rowNum++);
            row.createCell(0).setCellValue(stdfile);
            row.createCell(1).setCellValue(mistakeCount);
            row.createCell(2).setCellValue(totalMarks);
            row.createCell(3).setCellValue(extraWordCount);
            row.createCell(4).setCellValue(missingWordCount);
            row.createCell(5).setCellValue(wrongWordCount);
            row.createCell(6).setCellValue(allMistakes.toString()); // all mistakes in last column

            System.out.println("✅ Processed: " + stdfile);
        }

        // Save Excel summary
        FileOutputStream fosExcel = new FileOutputStream(
                new File(studentFolder.getParentFile(), "StudentMistakesSummary.xlsx"));
        summaryWorkbook.write(fosExcel);
        fosExcel.close();
        summaryWorkbook.close();

        System.out.println("✅ All files processed. Summary Excel created!");
    }

    private static String extractFullText(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis);
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph p : doc.getParagraphs()) {
            sb.append(p.getText()).append(" ");
        }
        fis.close();
        return sb.toString().replaceAll("\\r?\\n", " ").replaceAll("\\s+", " ").trim();
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
}
