package Test_Package;
import org.apache.poi.xwpf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
public class CreateDocumentDocxFile {


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

            // new document for output
            XWPFDocument outDoc = new XWPFDocument();
            XWPFParagraph para = outDoc.createParagraph();
            XWPFRun run = para.createRun();

            // compare word by word
            while (i < mWords.length || j < sWords.length) {
                String refWordRaw = i < mWords.length ? mWords[i] : null;
                String stuWordRaw = j < sWords.length ? sWords[j] : null;

                if (refWordRaw != null && stuWordRaw != null && refWordRaw.equals(stuWordRaw)) {
                    // correct word
                    run.setText(stuWordRaw + " ");
                    i++;
                    j++;
                } else if (refWordRaw != null && stuWordRaw != null) {
                    if (j + 1 < sWords.length && refWordRaw.equals(sWords[j + 1])) {
                        // extra word
                        XWPFRun r = para.createRun();
                        r.setText("["+ stuWordRaw + " "+"]");
                        r.setColor("FF0000"); // red
                        extraWordCount++;
                        j++;
                    } else if (i + 1 < mWords.length && mWords[i + 1].equals(stuWordRaw)) {
                        // missing word
                        XWPFRun r = para.createRun();
                        r.setText("["+refWordRaw + " "+"]");
                        r.setColor("FF0000");
                        missingWordCount++;
                        i++;
                    } else {
                        // wrong word
                        XWPFRun r = para.createRun();
                        r.setText("["+refWordRaw+" / "+stuWordRaw + " ]");
                        r.setColor("FF0000");
                        wrongWordCount++;
                        i++;
                        j++;
                    }
                } else {
                    if (refWordRaw != null) {
                        // missing word
                        XWPFRun r = para.createRun();
                        r.setText("["+refWordRaw + " ]");
                        r.setColor("FF0000");
                        missingWordCount++;
                        i++;
                    }
                    if (stuWordRaw != null) {
                        // extra word
                        XWPFRun r = para.createRun();
                        r.setText("["+stuWordRaw + " ]");
                        r.setColor("FF0000");
                        extraWordCount++;
                        j++;
                    }
                }
            }

            // calculate marks
            int MistakeCount = (missingWordCount + extraWordCount + wrongWordCount);
            if (MistakeCount <= 40) {
                Total_Marks = 40 - MistakeCount;
            }

            // add summary at bottom
            XWPFParagraph summaryPara = outDoc.createParagraph();
            summaryPara.setSpacingBefore(200);
            XWPFRun summaryRun = summaryPara.createRun();
            summaryRun.setText("\n--- Mistake Summary ---\n");
            summaryRun.setBold(true);
            summaryRun.addCarriageReturn();
            summaryRun = summaryPara.createRun();
            summaryRun.setText("Missing Words : " + missingWordCount + "\n");summaryRun.addCarriageReturn();
            summaryRun.setText("Extra Words   : " + extraWordCount + "\n");summaryRun.addCarriageReturn();
            summaryRun.setText("Wrong Words   : " + wrongWordCount + "\n");summaryRun.addCarriageReturn();
            summaryRun.setText("Total Marks   : " + Total_Marks + "\n");

            // save output file
            FileOutputStream fos = new FileOutputStream("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\marking\\Result_Highlighted.docx");
            outDoc.write(fos);
            fos.close();
            outDoc.close();

            System.out.println("✅ Output file created with highlights & summary!");
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

            // normalize
            return sb.toString()
                    .replaceAll("\\r?\\n", " ")
                    .replaceAll("\\s+", " ")
                    .trim();
        }
    }



