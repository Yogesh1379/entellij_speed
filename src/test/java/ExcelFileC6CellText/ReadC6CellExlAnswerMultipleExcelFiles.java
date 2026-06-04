package ExcelFileC6CellText;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadC6CellExlAnswerMultipleExcelFiles {

    public static void main(String[] args) throws IOException {
        String folderPath = "F:\\GCCTBC-APR 2026\\Mar_3040_AnswerFiles_Apr_2026"; // Folder containing all student Excel files
        String outputFilePath = "F:\\GCCTBC-APR 2026\\Wrong compare check\\marathi\\Excel wrong compare.xlsx"; // Output Excel file

        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx") || name.endsWith(".xls"));

        if (files == null || files.length == 0) {
            System.out.println("No Excel files found in folder!");
            return;
        }
        Workbook outputWorkbook;


            Sheet outputSheet;
            File outputFile = new File(outputFilePath);
            if (outputFile.exists()) {
                FileInputStream outFis = new FileInputStream(outputFile);
                outputWorkbook = new XSSFWorkbook(outFis);
                outputSheet = outputWorkbook.getSheetAt(0);
                outFis.close();
            } else {
                outputWorkbook = new XSSFWorkbook();
                outputSheet = outputWorkbook.createSheet("Sheet1");
            }

            int nextRowNum = outputSheet.getLastRowNum() + 1;

            for (File file : files) {
                try {
//                if(file.isFile()&&file.getName().startsWith("ExcelAnswer") && file.getName().endsWith(".xlsx")){
                    System.out.println(file.getName());
                    FileInputStream fis = new FileInputStream(file);
                    Workbook workbook = new XSSFWorkbook(fis);
                    Sheet sheet = workbook.getSheetAt(0); // Read first sheet

                    Row row = sheet.getRow(6); // C6 -> row 5 (0-based)
                    String cellValue = "";
                    if (row != null) {
                        Cell cell = row.getCell(2); // Column C -> index 2
                        if (cell != null) {
                            cellValue = cell.toString();
                        }
                    }

                    // Write file name and C6 value in output Excel
                    Row newRow = outputSheet.createRow(nextRowNum++);
                    newRow.createCell(0).setCellValue(file.getName());
                    newRow.createCell(1).setCellValue(cellValue);

                    workbook.close();
                    fis.close();
                } catch (Exception e) {

                    Row newRow = outputSheet.createRow(nextRowNum++);
                    newRow.createCell(0).setCellValue(file.getName());
                    newRow.createCell(1).setCellValue(e.getMessage());

                }
            }


            // Save the output file
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            outputWorkbook.write(fos);
            fos.close();
            outputWorkbook.close();

            System.out.println("All student files processed successfully!");


        }
    }

