package Statement_AnswerFile_Sheet2;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;

import java.io.*;

public class Sheet2FileName {
    public static void main(String[] args) throws IOException {


        IOUtils.setByteArrayMaxOverride(200_000_000);

        File inputFolder = new File("F:\\GCCTBC-APR 2026\\Mar_3040_AnswerFiles_Apr_2026\\");
        File[] exlFiles = inputFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));

        if (exlFiles == null || exlFiles.length == 0) {
            System.out.println("No files found");
            return;
        }

        Workbook summaryWorkbook = new XSSFWorkbook();
        Sheet summarySheet = summaryWorkbook.createSheet("File name");

        Row header = summarySheet.createRow(0);
        header.createCell(0).setCellValue("File name");
        header.createCell(1).setCellValue("Error");

        int rownum = 1;

        for (File excelFile : exlFiles) {
            System.out.println(excelFile.getName());
            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook wb = new XSSFWorkbook(fis)) {

                System.out.println(excelFile.getName());

                if (wb.getNumberOfSheets() == 2) {
                    Row row = summarySheet.createRow(rownum++);
                    row.createCell(0).setCellValue(excelFile.getName());
                }

            } catch (Exception e) {
                System.out.println("Error in file: " + excelFile.getName());

                Row row = summarySheet.createRow(rownum++);
                row.createCell(0).setCellValue( excelFile.getName());
                row.createCell(1).setCellValue( e.getMessage());
            }
        }

        try (FileOutputStream fosExcel = new FileOutputStream(
                "F:\\GCCTBC-APR 2026\\Sheet2 files\\Sheet2 Mar stmt file names.xlsx")) {

            summaryWorkbook.write(fosExcel);
        }

        summaryWorkbook.close();
    }
}