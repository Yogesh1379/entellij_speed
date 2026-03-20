package Test_Package;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Marathi_EnterAndTabCount {

static int FormatingMistake=0;
    static int spaceCount =0;
 static   int tabCount =0;
  static  int maxBlankSequence = 0;


        public static void main(String[] args) throws IOException {

//          File folder=  new File("C:\\Users\\User\\Desktop\\New folder (3)");
////          for output
//            File compareDir = new File(folder.getParentFile(), folder.getName() + "_Compared");
//            if (!compareDir.exists()) compareDir.mkdirs();
//
//            File[] slist = folder.listFiles((dir,name)->name.toLowerCase().endsWith(".docx"));
//            if(slist==null)
//            {
//                System.out.println("folder is empty");
//             return;
//            }
//
//            for(File studentfile1:slist)
//            {
//                String filename = studentfile1.getName();
//                String[] part = filename.split("_");
//                String seatno = part[1];
//                String batchname = part[2];
////        	String section = parts[0];
//                System.out.println(seatno);
//                if (seatno.length()>=10) {
//                    String course = seatno.substring(4, 6);
//                    int course1 = 0;
//                    if (course.equals("15")) {
////        			nteger.parseInt(course);
//                        course1 = 1;
//                    } else if (course.equals("16")) {
////        			System.out.println("16");
//                        course1 = 2;
//                    }
//                }
//            }

            File modelfile = new File("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\Eng 40 Speed 4.docx");
            File studentfile = new File("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\SpeedAnswer_8301160268_104_10-10-2025_16-21-49.docx");

            // full continuous text
            String modelText = extractFullText(modelfile);
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

            // compare word by word
            while (i < mWords.length || j < sWords.length) {
                String refWordRaw = i < mWords.length ? mWords[i] : null;
                String stuWordRaw = j < sWords.length ? sWords[j] : null;

                if (refWordRaw != null && stuWordRaw != null && refWordRaw.equals(stuWordRaw)) {
                    // correct word
                    XWPFRun r = para.createRun();
                    r.setText(stuWordRaw + " ");
                    i++;
                    j++;
                } else if (refWordRaw != null && stuWordRaw != null) {
                    if (j + 1 < sWords.length && refWordRaw.equals(sWords[j + 1])) {
                        // extra word
                        XWPFRun r = para.createRun();
                        r.setText("[ " + stuWordRaw + "] ");
                        r.setColor("0000FF");//blue
                        extraWordCount++;
                        j++;
                    } else if (i + 1 < mWords.length && mWords[i + 1].equals(stuWordRaw)) {
                        // missing word
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWordRaw + " ] ");
                        r.setColor("008000");//green
                        missingWordCount++;
                        i++;
                    } else {
                        // wrong word
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWordRaw + " / " + stuWordRaw + "] ");
                        r.setColor("FF0000");
                        wrongWordCount++;
                        i++;
                        j++;
                    }
                } else {
                    if (refWordRaw != null) {
                        // missing word
                        XWPFRun r = para.createRun();
                        r.setText("[" + refWordRaw + " ] ");
                        r.setColor("008000");//green
                        missingWordCount++;
                        i++;
                    }
                    if (stuWordRaw != null) {
                        // extra word
                        XWPFRun r = para.createRun();
                        r.setText("[ " + stuWordRaw + "] ");
                        r.setColor("0000FF");//blue
                        extraWordCount++;
                        j++;
                    }
                }
            }

            // calculate marks
            FormatingMistake=maxBlankSequence+tabCount+spaceCount;
            int MistakeCount = (missingWordCount + extraWordCount + wrongWordCount+FormatingMistake);
            if (MistakeCount <= 40) {
                Total_Marks = 40 - MistakeCount;
            }

            // add summary at bottom
            XWPFParagraph summaryPara = outDoc.createParagraph();
            summaryPara.setSpacingBefore(200);
            XWPFRun summaryRun = summaryPara.createRun();
            summaryRun.setText("\n--- Mistake Summary ---\n");
//            summaryRun.setBold(true);
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

            // save output file
            FileOutputStream fos = new FileOutputStream("F:\\Desktop backup 10-10-2025\\New folder\\New folder\\Result_Highlighted.docx");
            outDoc.write(fos);
            fos.close();
            outDoc.close();

            System.out.println("✅ Output file created with inline highlights & summary!");
        }

        // extract full normalized text from docx
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
// --- Count before normalization ---
             tabCount = countMatches(rawText, "\t{2,}");           // 3 or more tabs

             spaceCount = countMatches(rawText, " {2,}");          // 3 or more spaces

            System.out.println("Continuous Tabs (3+): " + tabCount);

            System.out.println("Continuous Spaces (3+): " + spaceCount);

            // normalize
            return sb.toString()
                    .replaceAll("\\r?\\n", " ")
                    .replaceAll("\\s+", " ")
                    .trim();
        }
    private static int countMatches(String text, String regex) {
        java.util.regex.Matcher m = java.util.regex.Pattern.compile(regex).matcher(text);
        int count = 0;
        while (m.find()) {
            count++;
        }
        return count;
    }
    }




