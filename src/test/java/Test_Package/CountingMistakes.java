package Test_Package;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CountingMistakes {

        public static void main(String[] args) throws IOException {

            File modelfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\English_30_40_Quetion_paper\\All_Subjective_Eng30\\Eng30 Speed 1.docx");
            File studentfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\marking\\New folder\\SpeedAnswer_1101150017_101_6-18-2025_9-16-55.docx");

            List<String> modelPara = extractFormattedtext(modelfile);
            List<String> studentpara = extractFormattedtext(studentfile);

            int stdInx = 0;
            int refIndex = 0;

            int Total_Marks=0;

            // mistake counters
            int missingWordCount = 0;
            int extraWordCount = 0;
            int wrongWordCount = 0;
            int missingParaCount = 0;

            // iterate both lists parallel
            while (stdInx < studentpara.size() && refIndex < modelPara.size()) {
                String stdLine = studentpara.get(stdInx);
                String modelLine = modelPara.get(refIndex);

                String[] sWords = stdLine.split("\\s+");
                String[] mWords = modelLine.split("\\s+");

                int i = 0, j = 0;

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

                stdInx++;
                refIndex++;
            }

            // remaining missing paragraphs
                while (refIndex < modelPara.size()) {
                String missingPara = modelPara.get(refIndex);
                System.out.println("Missing paragraph :- " + missingPara);
                missingParaCount++;

                if (missingParaCount >=1) {
                    String[] words = missingPara.split("\\s+");
                    missingWordCount = missingWordCount+words.length;
                }

                refIndex++;
            }

            // print final counts
            System.out.println("\n--- Mistake Summary ---");
            System.out.println("Missing Words : " + missingWordCount);
            System.out.println("Extra Words   : " + extraWordCount);
            System.out.println("Wrong Words   : " + wrongWordCount);
            System.out.println("Missing Paragraphs : " + missingParaCount);
            int MistakeCount=(missingWordCount+extraWordCount+wrongWordCount);
            if(MistakeCount<=40){
                Total_Marks=40-MistakeCount;
            }
            System.out.println("Total Marks : " + Total_Marks);
        }

        // extract paragraphs from docx
        private static List<String> extractFormattedtext(File file) throws IOException {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument doc = new XWPFDocument(fis);
            List<String> paragraphs = new ArrayList<>();

            for (XWPFParagraph p : doc.getParagraphs()) {
                String text = p.getText().replaceAll("\\s+", " ").trim();
                if (!text.isEmpty()) {
                    paragraphs.add(text);
                }
            }

            fis.close();
            return paragraphs;
        }
    }


