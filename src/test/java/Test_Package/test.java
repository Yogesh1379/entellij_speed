package Test_Package;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class test {
    public static void main(String[] args) throws IOException {

        File modelfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\allocation\\English_30_40_Quetion_paper\\All_Subjective_Eng30\\Eng30 Speed 1.docx");
        File studentfile = new File("F:\\desktop backup 5 july 25\\GCC TBC MAY 2025 AUTO\\marking\\New folder\\SpeedAnswer_1101150001_101_6-18-2025_9-17-10.docx");

        List<String> modelPara = extractFormattedtext(modelfile);
        List<String> studentpara = extractFormattedtext(studentfile);

        int stdInx = 0;
        int refIndex = 0;

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
                        j++;
                    } else if (i + 1 < mWords.length && mWords[i + 1].equals(stuWordRaw)) {
                        System.out.println("Missing word :- " + refWordRaw);
                        i++;
                    } else {
                        System.out.println("Wrong word :- " + stuWordRaw + " : " + refWordRaw);
                        i++;
                        j++;
                    }
                } else {
                    if (refWordRaw != null) {
                        System.out.println("Missing word :- " + refWordRaw);
                        i++;
                    }
                    if (stuWordRaw != null) {
                        System.out.println("Extra word :- " + stuWordRaw);
                        j++;
                    }
                }
            }

            stdInx++;
            refIndex++;
        }
        while (refIndex < modelPara.size()) {
            System.out.println("Missing paragraph :- " + modelPara.get(refIndex));
            refIndex++;
        }

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
