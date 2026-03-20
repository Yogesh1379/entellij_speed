package Test_Package;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.File;
import java.io.IOException;
public class colorTextExtract {



        public static void main(String[] args) throws IOException {
            FileInputStream fis = new FileInputStream(new File("C:\\Users\\User\\Desktop\\ResultFiles\\Diff_SpeedAnswer_1514150103_104_18-06-2025_15-36-52.docx"));
            XWPFDocument document = new XWPFDocument(fis);

            for (XWPFParagraph para : document.getParagraphs()) {
                for (XWPFRun run : para.getRuns()) {
                    String text = run.text();
                    String color = run.getColor();   // direct font color (hex)
                    String highlight = run.getTextHightlightColor().toString(); // highlight color

                    if (color == null) {
                        CTRPr   prop = run.getCTR().getRPr();
                        if (prop != null && prop.addNewBdr().isSetColor()) {
                            color = prop.getColorList().toString();
                        } else {
                            color = "000000"; // default black
                        }
                    }

                    System.out.println("Text: " + text +
                            " | FontColor: " + color +
                            " | Highlight: " + highlight);
                }
            }

            document.close();
            fis.close();
        }
    }
