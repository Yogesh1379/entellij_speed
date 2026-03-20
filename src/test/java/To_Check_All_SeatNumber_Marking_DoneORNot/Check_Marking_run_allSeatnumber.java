package To_Check_All_SeatNumber_Marking_DoneORNot;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.HashSet;

public class Check_Marking_run_allSeatnumber {

    static public void main(String[] args) throws IOException {
        File originalfile= new File("F:\\GCC  TBC December 2025\\Marking\\Eng40markingfiles.xlsx");// file which is given to marking with que name and std ans file
        File markingFile=new File("F:\\GCC  TBC December 2025\\Marking\\Eng 40\\SpeedMarkingWithMarksFinalEng30_40.xlsx"); //merge marking data
        DataFormatter df = new DataFormatter();
        HashSet<Object> filenameset = new HashSet<>();
        FileInputStream fisoriginal=new FileInputStream(originalfile);
        XSSFWorkbook workbook = new XSSFWorkbook(fisoriginal);
        Sheet  originalsheet=workbook.getSheetAt(1);
        FileInputStream fismark = new FileInputStream(markingFile);
        XSSFWorkbook Markworkbook = new XSSFWorkbook(fismark);
        XSSFSheet markSheet = Markworkbook.getSheetAt(1);
        // CREATE RESULT WORKBOOK
        Workbook resultWorkbook = new XSSFWorkbook();

        Sheet matchSheet = resultWorkbook.createSheet("Match Result");
        Sheet mismatchSheet = resultWorkbook.createSheet("Mismatch Result");

        Row matchHeader = matchSheet.createRow(0);
        matchHeader.createCell(0).setCellValue("Matched Files");

        Row mismatchHeader = mismatchSheet.createRow(0);
        mismatchHeader.createCell(0).setCellValue("Mis Matched Files");

        int matchRow = 1;
        int notFoundRow = 1;


        for(int i=1;i<=markSheet.getLastRowNum();i++)
        {
            Row row=markSheet.getRow(i);
           Cell cell= row.getCell(1);  //have to change for lrt spd stmt
            String cellValue = df.formatCellValue(cell).replaceAll("Diff_","");
System.out.println(cellValue);
            filenameset.add(cellValue);
        }

        for(int i=1;i<=originalsheet.getLastRowNum();i++)
        {
            Row row=originalsheet.getRow(i);
            Cell cell= row.getCell(4);  //have to change for lrt spd stmt
            String cellValue = df.formatCellValue(cell);
            if(filenameset.contains(cellValue)){
                Row r = matchSheet.getRow(matchRow);
                if (r == null) r = matchSheet.createRow(matchRow);
                r.createCell(0).setCellValue(cellValue);
                matchRow++;
                System.out.println("match :- "+cellValue);
            }
            else{
                Row r = mismatchSheet.getRow(notFoundRow);
                if (r == null) r = mismatchSheet.createRow(notFoundRow);
                r.createCell(0).setCellValue(cellValue);
                notFoundRow++;
                System.out.println("notfount :- "+cellValue);
            }
        }

        // SAVE FILE
        FileOutputStream fos = new FileOutputStream(
                "F:\\GCC  TBC December 2025\\Marking\\Marking remaining eng40\\spd Mismatch File with lms marking.xlsx");
        resultWorkbook.write(fos);

        fos.close();
        resultWorkbook.close();
        workbook.close();
        Markworkbook.close();

        System.out.println(" Result Excel created successfully");



    }
}
