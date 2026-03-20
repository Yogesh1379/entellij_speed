package Marathi_Ignore_Word;

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
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.commons.io.FilenameUtils.normalize;

public class Marathi_EnterAndTabCount3_lcs_diffwith_Ignore_Word {
    static int FormatingMistake = 0;
    static int spaceCount = 0;
    static int tabCount = 0;
    static int maxBlankSequence = 0;

    public static void main(String[] args) throws IOException {
        //  IGNORE EXCEL PATH
        String ignoreExcelPath =
                "F:\\GCC  TBC December 2025\\Marathi Ignore Word\\GCC TBC Mar30_40 Hindi 30_40 Speed ignore word_Final - Copy.xlsx";


        // Folder containing student files
        File studentFolder = new File("F:\\GCC  TBC December 2025\\Mararthi bucket\\std ans file\\MAR 40 ZERO CASES\\New folder");
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
        header.createCell(11).setCellValue("Ignore Word");
        header.createCell(12).setCellValue("Ignore Count");
        header.createCell(13).setCellValue("Batch");

        int rowNum = 1;

        // Load allocation Excel only once
        FileInputStream fisAlloc = new FileInputStream(
                "F:\\GCC  TBC December 2025\\Question\\allocation\\marathi\\SUbjectiveMarathiHin.xlsx");
        Workbook allocWorkbook = new XSSFWorkbook(fisAlloc);
        Sheet allocSheet = allocWorkbook.getSheetAt(0);
        // Load Ignore Excel ONCE (Important for speed)
        FileInputStream ignoreFis = new FileInputStream(ignoreExcelPath);
        XSSFWorkbook ignoreWorkbook = new XSSFWorkbook(ignoreFis);

        Map<String, String> ignoreMap = new HashMap<>();

        for (File studentFile : docxFiles) {
            String stdfile = studentFile.getName();
            String[] parts = stdfile.split("_");
            if (parts.length < 3) continue;

            String seatno = parts[1];
            String batchname = parts[2];
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


            String sheetName = getSheetName(course1);
            Sheet sheet = ignoreWorkbook.getSheet(sheetName);

            // FAST SEARCH

            for (Row row : sheet) {

                if (row.getRowNum() == 0) continue;

                String batch = getCellValue(row.getCell(1));
                if (!batch.equalsIgnoreCase(batchname)) continue;

                String paperWord = normalizeWord(getCellValue(row.getCell(2)));
                String studentWord = normalizeWord(getCellValue(row.getCell(3)));

                if (!paperWord.isEmpty() && !studentWord.isEmpty()) {
                    ignoreMap.put(paperWord, studentWord);
                }


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
                            (fileCandidate.startsWith("Mar 40 Speed") && course1 == 4) ||
                            (fileCandidate.startsWith("Hin30 Speed") && course1 == 5) ||
                            (fileCandidate.startsWith("Hin 40 Speed") && course1 == 6)) {
                        modelFile = new File(
                                "F:\\GCC  TBC December 2025\\Question\\allocation\\marathi\\All\\" + fileCandidate);
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

            // Compare word by word using LCS-based alignment and generate highlighted DOCX
            XWPFDocument outDoc = new XWPFDocument();
            XWPFParagraph para = outDoc.createParagraph();

            String[] mWords = modelText.split("\\s+");
            String[] sWords = studentText.split("\\s+");

            // Use LCS to align sequences so continuous missing/extra segments are handled correctly
            AlignmentResult align = computeAlignment(mWords, sWords);

            int missingWordCount = 0, extraWordCount = 0, wrongWordCount = 0, ignoreWordCount=0;
            StringBuilder allMistakes = new StringBuilder();

            // Walk alignment and create colored runs

            StringBuilder sb = new StringBuilder();
            for (AlignedToken token : align.alignedTokens) {
                XWPFRun r = para.createRun();
                if (token.type == TokenType.MATCH) {
                    r.setText(token.student != null ? token.student + " " : token.model + " ");
                } else if (token.type == TokenType.MISSING) {
                    // model word(s) missing in student
                    for (String mw : token.modelSequence) {
                        XWPFRun rr = para.createRun();
                        rr.setText("[" + mw + " ] ");
                        rr.setColor("008000");
                    }
                    missingWordCount += token.modelSequence.size();
                    allMistakes.append("[ Missing: ").append(String.join(" ", token.modelSequence)).append(" ]; ");
                } else if (token.type == TokenType.EXTRA) {
                    // student has extra word(s)
                    for (String sw : token.studentSequence) {
                        XWPFRun rr = para.createRun();
                        rr.setText("[ " + sw + "] ");
                        rr.setColor("0000FF");
                    }
                    extraWordCount += token.studentSequence.size();
                    allMistakes.append("[ Extra: ").append(String.join(" ", token.studentSequence)).append(" ]; ");
                } else if (token.type == TokenType.SUBSTITUTION) {
//            ============*****   ignore word  **** ===================


// inside SUBSTITUTION:
                    String normModel = normalizeWord(token.model);
                    String normStudent = normalizeWord(token.student);

                    if (ignoreMap.containsKey(normModel)) {

                        String expectedStudentWord = ignoreMap.get(normModel);

                        if (normStudent.equals(expectedStudentWord)) {

                            // This is valid ignore pair
                            sb.append(normStudent).append(", ");
                          ignoreWordCount++;
                            // Skip counting this as wrong word
//                            continue;
                        }
                    }

                    // substitution / wrong word(s)
                    XWPFRun rr = para.createRun();
                    rr.setText("[ " + token.model + " / " + token.student + "] ");
                    rr.setColor("FF0000");
                    wrongWordCount++;
                    allMistakes.append("[ Wrong: ").append(token.student).append(" | ").append(token.model).append(" ]; ");
                }
            }

            // Total formatting mistakes
            FormatingMistake = maxBlankSequence + tabCount + spaceCount;
            int totalMistakes = missingWordCount + extraWordCount + wrongWordCount + FormatingMistake-ignoreWordCount;
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
            summaryRun.setText("Ignore word Count: " + ignoreWordCount);
            summaryRun.addCarriageReturn();
            summaryRun.setText("Ignore word: " + sb.toString());
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

            row.createCell(11).setCellValue(sb.toString());
            row.createCell(12).setCellValue(ignoreWordCount);
            row.createCell(13).setCellValue(batchname);

            System.out.println(stdfile);
        }

        // Save Excel summary
        FileOutputStream fosExcel = new FileOutputStream(
                new File(studentFolder.getParentFile(), "Marathi Speed MistakesSummary.xlsx"));
        summaryWorkbook.write(fosExcel);
        fosExcel.close();
        summaryWorkbook.close();
        allocWorkbook.close();
        fisAlloc.close();

        System.out.println(" All files processed. Summary Excel created!");
    }

    // ----------------- NEW: LCS-based alignment -----------------
    private static AlignmentResult computeAlignment(String[] modelWords, String[] studentWords) {
        int n = modelWords.length;
        int m = studentWords.length;
        String[] normM = new String[n];
        String[] normS = new String[m];
        for (int i = 0; i < n; i++) normM[i] = normalizeWord(modelWords[i]);
        for (int j = 0; j < m; j++) normS[j] = normalizeWord(studentWords[j]);

        int[][] dp = new int[n + 1][m + 1];
        for (int i = 1; i <= n; i++) {
            for (int j = 1; j <= m; j++) {
                if (normM[i - 1].equals(normS[j - 1])) dp[i][j] = dp[i - 1][j - 1] + 1;
                else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
            }
        }

        // Backtrack to get LCS alignment positions
        List<Integer> mPos = new ArrayList<>();
        List<Integer> sPos = new ArrayList<>();
        int i = n, j = m;
        while (i > 0 && j > 0) {
            if (normM[i - 1].equals(normS[j - 1])) {
                mPos.add(i - 1);
                sPos.add(j - 1);
                i--; j--;
            } else if (dp[i - 1][j] >= dp[i][j - 1]) i--;
            else j--;
        }
        Collections.reverse(mPos);
        Collections.reverse(sPos);

        // Build aligned token list by walking over the sequences and LCS matches
        List<AlignedToken> aligned = new ArrayList<>();
        int mi = 0, si = 0;
        int k = 0; // index over matches
        while (mi < n || si < m) {
            int nextM = (k < mPos.size()) ? mPos.get(k) : n;
            int nextS = (k < sPos.size()) ? sPos.get(k) : m;

            // Segments before the next LCS match
            int modelSegmentLen = Math.max(0, nextM - mi);
            int studentSegmentLen = Math.max(0, nextS - si);

            if (modelSegmentLen > 0 || studentSegmentLen > 0) {
                // If both have content, try to treat as substitutions when possible
                int pairLen = Math.min(modelSegmentLen, studentSegmentLen);
                for (int p = 0; p < pairLen; p++) {
                    String mw = modelWords[mi + p];
                    String sw = studentWords[si + p];
                    if (normalizeWord(mw).equals(normalizeWord(sw))) {
                        aligned.add(AlignedToken.match(mw, sw));
                    } else {
                        aligned.add(AlignedToken.substitution(mw, sw));
                    }
                }
                // Remaining model words (if any) are missing
                if (modelSegmentLen > pairLen) {
                    aligned.add(AlignedToken.missing(Arrays.copyOfRange(modelWords, mi + pairLen, nextM)));
                }
                // Remaining student words (if any) are extra
                if (studentSegmentLen > pairLen) {
                    aligned.add(AlignedToken.extra(Arrays.copyOfRange(studentWords, si + pairLen, nextS)));
                }

                mi = nextM;
                si = nextS;
            }

            // If there's a match at nextM/nextS
            if (k < mPos.size()) {
                String mw = modelWords[nextM];
                String sw = studentWords[nextS];
                // These should be equal because they came from LCS, but normalize again for safety
                if (normalizeWord(mw).equals(normalizeWord(sw))) {
                    aligned.add(AlignedToken.match(mw, sw));
                } else {
                    aligned.add(AlignedToken.substitution(mw, sw));
                }
                mi = nextM + 1;
                si = nextS + 1;
                k++;
            } else {
                // No more matches; anything remaining are either missing or extra
                if (mi < n) {
                    aligned.add(AlignedToken.missing(Arrays.copyOfRange(modelWords, mi, n)));
                    mi = n;
                }
                if (si < m) {
                    aligned.add(AlignedToken.extra(Arrays.copyOfRange(studentWords, si, m)));
                    si = m;
                }
            }
        }

        return new AlignmentResult(aligned);
    }

    public enum TokenType { MATCH, MISSING, EXTRA, SUBSTITUTION }

    public static class AlignedToken {
        public TokenType type;
        public String model; // single model word for match/subst
        public String student; // single student word for match/subst
        public List<String> modelSequence; // for missing
        public List<String> studentSequence; // for extra

        public static AlignedToken match(String m, String s) {
            AlignedToken t = new AlignedToken();
            t.type = TokenType.MATCH;
            t.model = m; t.student = s;
            return t;
        }

        public static AlignedToken missing(String[] seq) {
            AlignedToken t = new AlignedToken();
            t.type = TokenType.MISSING;
            t.modelSequence = Arrays.asList(seq);
            return t;
        }

        public static AlignedToken extra(String[] seq) {
            AlignedToken t = new AlignedToken();
            t.type = TokenType.EXTRA;
            t.studentSequence = Arrays.asList(seq);
            return t;
        }

        public static AlignedToken substitution(String m, String s) {
            AlignedToken t = new AlignedToken();
            t.type = TokenType.SUBSTITUTION;
            t.model = m; t.student = s;
            return t;
        }
    }

    public static class AlignmentResult {
        public List<AlignedToken> alignedTokens;
        public AlignmentResult(List<AlignedToken> a) { this.alignedTokens = a; }
    }

    // ----------------- END LCS-based alignment -----------------

    private static String extractFullText(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis);
        StringBuilder sb = new StringBuilder();

        int totalTabCount = 0;
        int totalSpaceCount = 0;

        int continuousBlankParas = 0;
        maxBlankSequence = 0;

        //  Get all paragraphs first
        var paragraphs = doc.getParagraphs();

        //  Find the index of the last *non-empty* paragraph (with visible text)
        int lastTextIndex = -1;
        for (int i = paragraphs.size() - 1; i >= 0; i--) {
            String t = paragraphs.get(i).getText();
            if (t != null && !t.trim().isEmpty()) {
                lastTextIndex = i;
                break;
            }
        }

        for (int idx = 0; idx < paragraphs.size(); idx++) {
            XWPFParagraph p = paragraphs.get(idx);
            String text = p.getText();

            // Skip trailing blank paragraphs (only those after the last text paragraph)
            if (idx > lastTextIndex) continue;

            // Count blank paragraphs (for Extra Enter)
            if (text.trim().isEmpty()) {
                continuousBlankParas++;
                if (continuousBlankParas >= 1) maxBlankSequence++;
            } else {
                continuousBlankParas = 0;
            }

            //  Ignore starting tabs/spaces before first visible character
            String textAfterLeading = text.replaceAll("^[\\t ]+", "");

            //  Count extra spaces/tabs only after text begins
            totalTabCount += countMatches(textAfterLeading, "\\t{1,}");

            // Count extra spaces beyond 2 continuous
            Matcher spaceMatcher = Pattern.compile(" {2,}").matcher(textAfterLeading);
            while (spaceMatcher.find()) {
                int len = spaceMatcher.end() - spaceMatcher.start();
                totalSpaceCount += (len - 2); // only spaces beyond 2 are extra
            }

            sb.append(text).append("\n ");
        }

        fis.close();

        //  Now ignore trailing tabs/spaces if they exist at document end
        String finalText = sb.toString();
        finalText = finalText.replaceAll("[\\t\\n\\r ]+$", ""); // remove trailing tabs/spaces/enters

        // Save counts globally
        tabCount = totalTabCount;
        spaceCount = totalSpaceCount;

        // Return clean text for comparison
        return finalText.replaceAll("\\r?\\n", " ").replaceAll("\\s+", " ").trim();
    }

    private static int countMatches(String text, String regex) {
        Matcher m = Pattern.compile(regex).matcher(text);
        int count = 0;
        while (m.find()) count++;
        return count;
    }

    public static String getCellValueAsString(Cell cell) {
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
    public static String normalizeWord(String word) {
        if (word == null) return null;

        // Replace curly and accented quotes with straight ones
        return word
                .replace("’", "'")   // curly apostrophe
                .replace("‘", "'")   // opening curly single quote
                .replace("´", "'")   // acute accent often used as apostrophe
                .replace("`", "'")   // grave accent
                .replace("“", "\"")  // opening curly double quote
                .replace("”", "\"") // closing curly double quote
                .trim();
    }

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

}
