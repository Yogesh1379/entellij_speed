package Import_Data_Creation;

import Send_Mail_Package.SendMailOf_ImportDataCreation;
import jakarta.mail.MessagingException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.testng.annotations.Test;

import java.io.*;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;

public class CreationImportData {
    @Test
    public static void importData_Creation() throws IOException, InvalidFormatException, MessagingException {
       File inputFile= new File("F:\\GCC TBC JULY\\STUDENT DATA\\New folder\\ENGLISH STUDENT DATA TOTAL106055.xlsx");
      String output="F:\\GCC TBC JULY\\STUDENT DATA\\New folder\\abc.xlsx";
       File outputFile=new File(output);
       FileInputStream fis=new FileInputStream(inputFile);
        org.apache.poi.ss.usermodel.Workbook wk =XSSFWorkbookFactory.createWorkbook(OPCPackage.open(fis));
        Sheet  sheet=wk.getSheetAt(0);
        int rc = sheet.getLastRowNum();
       DataFormatter df=new DataFormatter();

        Map<String, Integer> studentMap = new LinkedHashMap<>();
        Map<String, String> photoMap = new LinkedHashMap<>();
        XSSFWorkbook wb1=new XSSFWorkbook();
        XSSFSheet outSheet = wb1.createSheet("Import Data");
        Row h=outSheet.createRow(0);
        h.createCell(0).setCellValue("INSTITUTE_ID");
        h.createCell(1).setCellValue("STUDENT_FIRST_NAME");
        h.createCell(2).setCellValue("STUDENT_MIDDLE_NAME");
        h.createCell(3).setCellValue("STUDENT_LAST_NAME");
        h.createCell(4).setCellValue("MOTHER_NAME");
        h.createCell(5).setCellValue("EMAIL_ID");
        h.createCell(6).setCellValue("GENDER");
        h.createCell(7).setCellValue("MOBILE_NO");
        h.createCell(8).setCellValue("DATE_OF_BIRTH");
        h.createCell(9).setCellValue("EDUCATION");
        h.createCell(10).setCellValue("COURSE_SUBJECT");
        h.createCell(11).setCellValue("ADMISSION_DATE");
        h.createCell(12).setCellValue("HANDICAP");
        h.createCell(13).setCellValue("ADDRESS");
        h.createCell(14).setCellValue("IS_HELPER");
        h.createCell(15).setCellValue("HELPER_NAME");
        h.createCell(16).setCellValue("PHOTO");
        h.createCell(17).setCellValue("IS_REPEATER");
        h.createCell(18).setCellValue("FORM_SERIAL_NO");
        int serial = 1;
        int r=1;
       for(int i=1;i<=rc;i++){
         Row row=sheet.getRow(i);
         if(row==null) continue;
           String instId=df.formatCellValue(row.getCell(0)).trim();
           String firstName=df.formatCellValue(row.getCell(1)).trim();
           String middleName=df.formatCellValue(row.getCell(2)).trim();
           String lastName=df.formatCellValue(row.getCell(3)).trim();
           String motherName= df.formatCellValue(row.getCell(4)).trim();
           String email=df.formatCellValue(row.getCell(5)).trim();
           String Mobile=df.formatCellValue(row.getCell(6)).trim();
           String gender=df.formatCellValue(row.getCell(7)).trim();
           String handi= df.formatCellValue(row.getCell(8)).trim();
           String admission=df.formatCellValue(row.getCell(9)).trim();
           String DOB= df.formatCellValue(row.getCell(10)).trim();
           String education=df.formatCellValue(row.getCell(11)).trim();
           String address= df.formatCellValue(row.getCell(12)).trim();
           String courseSubject=df.formatCellValue(row.getCell(13)).trim();
           String photo=df.formatCellValue(row.getCell(21)).trim();
           String isRepeater=df.formatCellValue(row.getCell(22)).trim();

          System.out.println(i);

          Row rowOut=outSheet.createRow(r++);
           rowOut.createCell(0).setCellValue(instId);
           rowOut.createCell(1).setCellValue(firstName);
           rowOut.createCell(2).setCellValue(middleName);
           rowOut.createCell(3).setCellValue(lastName);
           rowOut.createCell(4).setCellValue(motherName);
           rowOut.createCell(5).setCellValue(email);
           rowOut.createCell(6).setCellValue(gender);
           int mo = Mobile.length();
           if(!(Mobile.isEmpty())&& mo==10){
               rowOut.createCell(7).setCellValue(Mobile);
           }
           else if(!Mobile.isEmpty() && mo<10 ){
           rowOut.createCell(7).setCellValue("9876543210");}

           if(DOB.equalsIgnoreCase("NULL")){
               rowOut.createCell(8).setCellValue("11-11-2111");}
           else {
               rowOut.createCell(8).setCellValue(DOB);
           }

           rowOut.createCell(9).setCellValue(education);
           if((courseSubject.equalsIgnoreCase("GCC TBC English 30 wpm"))||(courseSubject.equalsIgnoreCase("GCC TBC ENG 30"))){
               rowOut.createCell(10).setCellValue("English30");
           } else if ((courseSubject.equalsIgnoreCase("GCC TBC English 40 wpm"))||(courseSubject.equalsIgnoreCase("GCC TBC ENG 40"))) {
               rowOut.createCell(10).setCellValue("English40");
           } else if (courseSubject.equalsIgnoreCase("GCC TBC Marathi 30 wpm")) {
               rowOut.createCell(10).setCellValue("Marathi30");
           }else if (courseSubject.equalsIgnoreCase("GCC TBC Marathi 40 wpm")) {
               rowOut.createCell(10).setCellValue("Marathi40");
           }else if (courseSubject.equalsIgnoreCase("GCC TBC Hindi 30 wpm")) {
               rowOut.createCell(10).setCellValue("Hindi30");
           }else if (courseSubject.equalsIgnoreCase("GCC TBC Hindi 40 wpm"))  {
               rowOut.createCell(10).setCellValue("Hindi40");
           }


           rowOut.createCell(11).setCellValue(admission);
           rowOut.createCell(12).setCellValue(handi);
           rowOut.createCell(13).setCellValue(address);
           rowOut.createCell(14).setCellValue(0);
           rowOut.createCell(15).setCellValue("NA");

//           rowOut.createCell(16).setCellValue(photo);
           rowOut.createCell(17).setCellValue(isRepeater);

           String key = instId + "|" +firstName + "|" +middleName + "|" +lastName + "|" +motherName;

           if (!photoMap.containsKey(key)) {
               photoMap.put(key, photo);
           }

//  always use same photo for same student
           String finalPhoto = photoMap.get(key);

//           rowOut.createCell(16).setCellValue(finalPhoto+".JPG");
           rowOut.createCell(16).setCellValue(finalPhoto);

//           ******************************
           if (!studentMap.containsKey(key)) {
               studentMap.put(key, serial++);
           }
           int fsn = studentMap.get(key);
           rowOut.createCell(18).setCellValue(fsn);

       }
        FileOutputStream fos = new FileOutputStream(outputFile);
        wb1.write(fos);
        fos.close();
        wb1.close();
//        SendMailOf_ImportDataCreation.sedReports(output);
        System.out.println("EXCEL generated successfully!");

    }
}
