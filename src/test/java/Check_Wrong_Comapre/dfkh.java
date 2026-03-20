package Check_Wrong_Comapre;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

public class dfkh {

    public static void main(String[] args) throws Exception {

        File folder = new File("C:\\Users\\User\\Desktop\\New folder\\New folder");

        for (File file : folder.listFiles()) {

            if (!file.getName().endsWith(".docx")) continue;

            System.out.println("Processing: " + file.getName());

            XWPFDocument doc = new XWPFDocument(new FileInputStream(file));
            List<XWPFParagraph> paras = doc.getParagraphs();

            XWPFParagraph startPara = null;   // First non-empty paragraph → start
            XWPFParagraph endPara = null;     // Paragraph that matches email/web
            boolean bookmarkApplied = false;

            for (XWPFParagraph para : paras) {
                String text = para.getText().trim();
                if (text.isEmpty()) continue;

                // 1️⃣ First non-empty paragraph → start
                if (startPara == null) {
                    startPara = para;
                }

                String loText = text.toLowerCase();

                // 2️⃣ Check if paragraph matches email/web condition
                if (!bookmarkApplied &&
                        (loText.contains("e-mail") ||
                                loText.contains("email") ||
                                loText.contains("website") ||
                                loText.contains("web"))) {

                    endPara = para;

                    if (startPara != null) {
                        addBookmarkRange(startPara, endPara, "HeadingBookmark");
                        bookmarkApplied = true;
                        System.out.println("Bookmark added from first paragraph → matched paragraph.");
                    }
                }
            }

            // Save updated file
            FileOutputStream fos = new FileOutputStream(
                    "C:\\Users\\User\\Desktop\\New folder\\OUT_" + file.getName()
            );
            doc.write(fos);
            fos.close();
            doc.close();

            System.out.println("✔ File processed successfully.");
        }
    }

    public static void addBookmarkRange(XWPFParagraph startPara,
                                        XWPFParagraph endPara,
                                        String bookmarkName) {

        int id = (int) (Math.random() * 100000);

        // ---- Ensure start paragraph has at least one run ----
        if (startPara.getRuns().isEmpty()) {
            XWPFRun run = startPara.createRun();
            run.setText(startPara.getText());
            startPara.removeRun(0);
        }

        // ---- Bookmark Start ----
        CTBookmark bookmarkStart = startPara.getCTP().addNewBookmarkStart();
        bookmarkStart.setId(BigInteger.valueOf(id));
        bookmarkStart.setName(bookmarkName);

        // ---- Ensure end paragraph has at least one run ----
        if (endPara.getRuns().isEmpty()) {
            XWPFRun run = endPara.createRun();
            run.setText(endPara.getText());
            endPara.removeRun(0);
        }

        // ---- Bookmark End ----
        CTMarkupRange bookmarkEnd = endPara.getCTP().addNewBookmarkEnd();
        bookmarkEnd.setId(BigInteger.valueOf(id));
    }
}
