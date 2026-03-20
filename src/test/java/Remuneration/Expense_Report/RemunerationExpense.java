package Remuneration.Expense_Report;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class RemunerationExpense {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream(
                "C:\\Users\\User\\Desktop\\remunatuin\\Book1.xlsx"
        );
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rc = sheet.getLastRowNum();


        DataFormatter df = new DataFormatter();
        Workbook workbook1 = new XSSFWorkbook();
        Sheet sheet1 = workbook1.createSheet("TestSheet");
        Row header = sheet1.createRow(0);
        header.createCell(0).setCellValue("Center Code");
        header.createCell(1).setCellValue("Center Name");
        header.createCell(2).setCellValue("Max PC Count");
        header.createCell(3).setCellValue("It Teacher Count");
        header.createCell(4).setCellValue("It Teacher Amount");
        header.createCell(5).setCellValue("winner Tech");
        header.createCell(6).setCellValue("winner Tech Amount");
        header.createCell(7).setCellValue("Peon");
        header.createCell(8).setCellValue("Peon Amount");
        header.createCell(9).setCellValue("Center Head ");
        header.createCell(10).setCellValue("Center Head Amount");
        header.createCell(11).setCellValue("College Amount");
        header.createCell(12).setCellValue("Other");
        header.createCell(13).setCellValue("Exam Day");


        for(int i=1;i<=rc;i++){
            XSSFRow row = sheet.getRow(i);
            Row row1 = sheet1.createRow(i);
            String centreCode = df.formatCellValue(row.getCell(1));
            String centerName = df.formatCellValue(row.getCell(2));
            String engDay = df.formatCellValue(row.getCell(3));
            String engPC=df.formatCellValue(row.getCell(4));
            String engStd=df.formatCellValue(row.getCell(5));
            Integer EngSTD = Integer.valueOf(engStd);
            String marday=df.formatCellValue(row.getCell(6));
            String marPC=df.formatCellValue(row.getCell(7));
            String marStd=df.formatCellValue(row.getCell(8));
            Integer MarSTD=Integer.valueOf(marStd);
           String eng50std= df.formatCellValue(row.getCell(11));
           Integer Eng50STD= Integer.valueOf(eng50std);
           String eng60std= df.formatCellValue(row.getCell(14));
            Integer Eng60STD=Integer.valueOf(eng60std);
           String totalStd=df.formatCellValue(row.getCell(15));
            Integer TotatlSTD=Integer.valueOf(totalStd);
            String examDay=df.formatCellValue(row.getCell(16));
            String mockday=df.formatCellValue(row.getCell(17));
            String totalDays=df.formatCellValue(row.getCell(18));

            int maxDay= Math.max(Integer.valueOf(engDay),Integer.valueOf(marday));

            int maxPCCount= Math.max(Integer.valueOf(engPC),Integer.valueOf(marPC));
            int requiredmanPower=(int) Math.ceil(maxPCCount/30.0);
            int requiredmancenterHead= (int) Math.ceil(maxPCCount/50.0);
            row1.createCell(0).setCellValue(centreCode);
            row1.createCell(1).setCellValue(centerName);
            row1.createCell(2).setCellValue(maxPCCount);
            //IT teacher
            row1.createCell(3).setCellValue(requiredmanPower);
            //IT teacher Amount
            row1.createCell(4).setCellValue(Integer.valueOf(totalDays)*requiredmanPower*700);
//winner tech
            row1.createCell(5).setCellValue(requiredmanPower);
            row1.createCell(6).setCellValue(Integer.valueOf(totalDays)*requiredmanPower*450);

            // peon
            row1.createCell(7).setCellValue(requiredmanPower);
            row1.createCell(8).setCellValue(Integer.valueOf(examDay)*requiredmanPower*150);

            // center head
            row1.createCell(9).setCellValue(requiredmancenterHead);
            row1.createCell(10).setCellValue(Integer.valueOf(examDay)*requiredmancenterHead*700);

            //college
            if((EngSTD+MarSTD+Eng50STD+Eng60STD)==TotatlSTD){
                row1.createCell(11).setCellValue(TotatlSTD*50);
            }

            //other
            row1.createCell(12).setCellValue(requiredmanPower*500);

            row1.createCell(13).setCellValue(totalDays);

            FileOutputStream fileOut = new FileOutputStream(new File(
                   "C:\\Users\\User\\Desktop\\remunatuin\\Remunation Expense report.xlsx"
            ));
            // Write the workbook data to the file
            workbook1.write(fileOut);
        }
        System.out.println("file out" );
    }
}
