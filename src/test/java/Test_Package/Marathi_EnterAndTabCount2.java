package Test_Package;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;
import java.util.regex.*;

public class Marathi_EnterAndTabCount2 {

    static int FormatingMistake = 0;
    static int spaceCount = 0;
    static int tabCount = 0;
    static int maxBlankSequence = 0;

    public static void main(String[] args) throws IOException {

        File modelfile = new File("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\Eng 40 Speed 4.docx");
        File studentfile = new File("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\SpeedAnswer_8301160268_104_10-10-2025_16-21-49.docx");

        String modelText = extractFullText(modelfile);
        String studentText = extractFullText(studentfile);

        String[] mWords = modelText.isEmpty() ? new String[0] : modelText.split("\\s+");
        String[] sWords = studentText.isEmpty() ? new String[0] : studentText.split("\\s+");

        int i = 0, j = 0;
        int missingWordCount = 0;
        int extraWordCount = 0;
        int wrongWordCount = 0;
        int Total_Marks = 0;

        XWPFDocument outDoc = new XWPFDocument();
        XWPFParagraph para = outDoc.createParagraph();

        // dynamic window size (bounded)
        int windowSize = Math.max(4, Math.min(40, Math.abs(mWords.length - sWords.length) ));
        System.out.println("Dynamic window size used: " + windowSize);

        while (i < mWords.length || j < sWords.length) {

            // both exist and match exactly
            if (i < mWords.length && j < sWords.length && mWords[i].equals(sWords[j])) {
                XWPFRun r = para.createRun();
                r.setText(sWords[j] + " ");
                i++;
                j++;
                continue;
            }

            // both exist but mismatch -> first check similarity of current pair
            if (i < mWords.length && j < sWords.length) {
                String mwRaw = mWords[i];
                String swRaw = sWords[j];
                String mwNorm = normalizeWord(mwRaw);
                String swNorm = normalizeWord(swRaw);

                // if words are similar enough => treat as wrong/substitution
                if (isSimilar(mwNorm, swNorm)) {
                    XWPFRun r = para.createRun();
                    r.setText("[" + mWords[i] + " / " + sWords[j] + "] ");
                    r.setColor("FF0000"); // red = wrong
                    wrongWordCount++;
                    i++;
                    j++;
                    continue;
                }

                // If similarity check didn't trigger, do look-ahead to find alignment
                int lookAheadModel = -1, lookAheadStudent = -1;
                outer:
                for (int a = 1; a <= windowSize && i + a < mWords.length; a++) {
                    for (int b = 1; b <= windowSize && j + b < sWords.length; b++) {
                        if (mWords[i + a].equals(sWords[j + b])) {
                            lookAheadModel = i + a;
                            lookAheadStudent = j + b;
                            break outer;
                        }
                    }
                }

                if (lookAheadModel != -1) {
                    // If alignment resumes immediately (i+1 and j+1) treat as single wrong/substitution
                    if (lookAheadModel == i + 1 && lookAheadStudent == j + 1) {
                        XWPFRun r = para.createRun();
                        r.setText("[" + mWords[i] + " / " + sWords[j] + "] ");
                        r.setColor("FF0000"); // red = wrong
                        wrongWordCount++;
                        i++;
                        j++;
                    } else {
                        // alignment found farther ahead -> mark block of missing (model) then extra (student)
                        while (i < lookAheadModel) {
                            XWPFRun r = para.createRun();
                            r.setText("[" + mWords[i] + " ] ");
                            r.setColor("008000"); // green = missing
                            missingWordCount++;
                            i++;
                        }
                        while (j < lookAheadStudent) {
                            XWPFRun r = para.createRun();
                            r.setText("[ " + sWords[j] + "] ");
                            r.setColor("0000FF"); // blue = extra
                            extraWordCount++;
                            j++;
                        }
                    }
                    continue;
                } else {
                    // no alignment found within window -> treat current as wrong substitution
                    XWPFRun r = para.createRun();
                    r.setText("[" + mWords[i] + " / " + sWords[j] + "] ");
                    r.setColor("FF0000");
                    wrongWordCount++;
                    i++;
                    j++;
                    continue;
                }
            }

            // one side exhausted -> remaining are missing or extra
            if (i < mWords.length) {
                XWPFRun r = para.createRun();
                r.setText("[" + mWords[i] + " ] ");
                r.setColor("008000");
                missingWordCount++;
                i++;
            } else if (j < sWords.length) {
                XWPFRun r = para.createRun();
                r.setText("[ " + sWords[j] + "] ");
                r.setColor("0000FF");
                extraWordCount++;
                j++;
            }
        }

        // --- Calculate marks ---
        FormatingMistake = maxBlankSequence + tabCount + spaceCount;
        int MistakeCount = (missingWordCount + extraWordCount + wrongWordCount + FormatingMistake);
        if (MistakeCount <= 40) {
            Total_Marks = 40 - MistakeCount;
        }

        // --- Summary Section ---
        XWPFParagraph summaryPara = outDoc.createParagraph();
        summaryPara.setSpacingBefore(200);
        XWPFRun summaryRun = summaryPara.createRun();
        summaryRun.setText("\n--- Mistake Summary ---\n");
        summaryRun.addCarriageReturn();

        summaryRun.setText("Missing Words (Green) : " + missingWordCount);
        summaryRun.addCarriageReturn();

        summaryRun.setText("Extra Words  (Blue)   : " + extraWordCount);
        summaryRun.addCarriageReturn();

        summaryRun.setText("Continue 2 or more Enter   : " + maxBlankSequence);
        summaryRun.addCarriageReturn();

        summaryRun.setText("Continue 3 or more Tab : " + tabCount);
        summaryRun.addCarriageReturn();

        summaryRun.setText("Continue 3 or more Spaces : " + spaceCount);
        summaryRun.addCarriageReturn();

        summaryRun.setText("Wrong Words  (Red)   : " + wrongWordCount);
        summaryRun.addCarriageReturn();

        summaryRun.addCarriageReturn();
        summaryRun.setBold(true);
        summaryRun.setText("Total Marks   : " + Total_Marks);
        summaryRun.addCarriageReturn();

        // --- Save output ---
        FileOutputStream fos = new FileOutputStream("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\Result_Highlighted.docx");
        outDoc.write(fos);
        fos.close();
        outDoc.close();

        System.out.println("✅ Output file created with improved missing-word handling & similarity detection!");
    }

    // --- Normalize a word for similarity checks: lowercase, trim punctuation at ends ---
    private static String normalizeWord(String w) {
        if (w == null) return "";
        // remove leading/trailing punctuation (keeps inner punctuation like apostrophe)
        return w.replaceAll("^\\p{Punct}+|\\p{Punct}+$", "").toLowerCase().trim();
    }

    // --- Decide if two normalized words are 'similar enough' to be a substitution ---
    private static boolean isSimilar(String a, String b) {
        if (a.isEmpty() && b.isEmpty()) return true;
        if (a.equals(b)) return true;

        // very short words: if one is single char and other is short, consider similar
        if (Math.max(a.length(), b.length()) <= 2) {
            return levenshtein(a, b) <= 1;
        }

        int dist = levenshtein(a, b);
        // absolute small distance OR small relative distance
        if (dist <= 2) return true;
        double rel = (double) dist / Math.max(1, Math.max(a.length(), b.length()));
        return rel <= 0.25; // up to 25% of chars can differ
    }

    // --- Levenshtein distance (classic DP) ---
    private static int levenshtein(String s1, String s2) {
        if (s1 == null) s1 = "";
        if (s2 == null) s2 = "";
        int n = s1.length();
        int m = s2.length();
        if (n == 0) return m;
        if (m == 0) return n;

        int[] prev = new int[m + 1];
        int[] cur = new int[m + 1];

        for (int j = 0; j <= m; j++) prev[j] = j;

        for (int i = 1; i <= n; i++) {
            cur[0] = i;
            for (int j = 1; j <= m; j++) {
                int cost = s1.charAt(i - 1) == s2.charAt(j - 1) ? 0 : 1;
                cur[j] = Math.min(Math.min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
            }
            // swap
            int[] tmp = prev;
            prev = cur;
            cur = tmp;
        }
        return prev[m];
    }

    // --- Extract and normalize text from docx ---
    private static String extractFullText(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis);

        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph p : doc.getParagraphs()) {
            sb.append(p.getText()).append("\n ");
        }

        fis.close();
        String rawText = sb.toString();
        int continuousBlankParas = 0;

        for (XWPFParagraph p : doc.getParagraphs()) {
            if (p.getText().trim().isEmpty()) {
                continuousBlankParas++;
                if (continuousBlankParas >= 1) {
                    maxBlankSequence++;
                }
            } else {
                continuousBlankParas = 0;
            }
        }

        System.out.println("Continuous blank paragraphs (2+): " + maxBlankSequence);

        // Count tabs and spaces
        tabCount = countMatches(rawText, "\t{2,}");
        spaceCount = countMatches(rawText, " {2,}");

        System.out.println("Continuous Tabs (3+): " + tabCount);
        System.out.println("Continuous Spaces (3+): " + spaceCount);

        // normalize: collapse whitespace but preserve simple word order
        return sb.toString()
                .replaceAll("\\r?\\n", " ")
                .replaceAll("\\s+", " ")
                .trim();
    }

    private static int countMatches(String text, String regex) {
        Matcher m = Pattern.compile(regex).matcher(text);
        int count = 0;
        while (m.find()) {
            count++;
        }
        return count;
    }
}
