package Test_Package;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ContinuesLineCount {

        public static void main(String[] args) throws IOException {

            File modelfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\English_30_40_Quetion_paper\\All_Subjective_Eng30\\Eng30 Speed 1.docx");
            File studentfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\marking\\New folder\\SpeedAnswer_1101150027_101_6-18-2025_9-17-01.docx");

            // full continuous text
            String modelText   = extractFullText(modelfile);
            String studentText = extractFullText(studentfile);

            // split into words
            String[] mWords = modelText.split("\\s+");
            String[] sWords = studentText.split("\\s+");

            int i = 0, j = 0;
            int missingWordCount = 0;
            int extraWordCount = 0;
            int wrongWordCount = 0;
            int Total_Marks = 0;

            // compare word by word
            while (i < mWords.length || j < sWords.length) {
                String refWordRaw = i < mWords.length ? mWords[i] : null;
                String stuWordRaw = j < sWords.length ? sWords[j] : null;

                if (refWordRaw != null && stuWordRaw != null && refWordRaw.equals(stuWordRaw)) {
                    // words match
                    i++;
                    j++;
                } else if (refWordRaw != null && stuWordRaw != null) {
                    if (j + 1 < sWords.length && refWordRaw.equals(sWords[j + 1])) {
                        System.out.println("Extra word :- " + stuWordRaw);
                        extraWordCount++;
                        j++;
                    } else if (i + 1 < mWords.length && mWords[i + 1].equals(stuWordRaw)) {
                        System.out.println("Missing word :- " + refWordRaw);
                        missingWordCount++;
                        i++;
                    } else {
                        System.out.println("Wrong word :- " + stuWordRaw + " : " + refWordRaw);
                        wrongWordCount++;
                        i++;
                        j++;
                    }
                } else {
                    if (refWordRaw != null) {
                        System.out.println("Missing word :- " + refWordRaw);
                        missingWordCount++;
                        i++;
                    }
                    if (stuWordRaw != null) {
                        System.out.println("Extra word :- " + stuWordRaw);
                        extraWordCount++;
                        j++;
                    }
                }
            }

            // print final counts
            System.out.println("\n--- Mistake Summary ---");
            System.out.println("Missing Words : " + missingWordCount);
            System.out.println("Extra Words   : " + extraWordCount);
            System.out.println("Wrong Words   : " + wrongWordCount);

            int MistakeCount = (missingWordCount + extraWordCount + wrongWordCount);
            if (MistakeCount <= 40) {
                Total_Marks = 40 - MistakeCount;
            }
            System.out.println("Total Marks : " + Total_Marks);
        }

        // extract full normalized text from docx
        private static String extractFullText(File file) throws IOException {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument doc = new XWPFDocument(fis);

            StringBuilder sb = new StringBuilder();
            for (XWPFParagraph p : doc.getParagraphs()) {
                sb.append(p.getText()).append(" ");
            }

            fis.close();

            // normalize -> enter ला space, multiple spaces एकाच space मध्ये
            return sb.toString()
                    .replaceAll("\\r?\\n", " ")
                    .replaceAll("\\s+", " ")
                    .trim();
        }
    }

