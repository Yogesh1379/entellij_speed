package Copy_File_With_SeatNumber;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;

public class Copy_All_Files_By_SeatNumber_Final {

    public static void main(String[] args) {

//    	//ENGLISH path
//    	String excelFilePath = "./excelfiles/stmt50.xlsx";
//    	String sourceFolderPath ="D:\\GCCTBCFinalBucketJanuary2026230126\\gcctbcEnglishfiles";
//    	String destinationFolderPath = "C:\\Users\\Project Head\\Desktop\\Recovery\\1";

        // MARATHI path
        String excelFilePath = "F:\\GCC  TBC December 2025\\Marking\\marathi 40\\Book1.xlsx";
        String sourceFolderPath = "F:\\GCC  TBC December 2025\\Mararthi bucket\\std ans file\\speed\\New folder";
        String destinationFolderPath = "F:\\GCC  TBC December 2025\\Mararthi bucket\\std ans file\\MAR 40 ZERO CASES\\New folder";

        try {
            // 1. Read seat numbers from Excel
            Set<String> seatNumbers = readSeatNumbersFromExcel(excelFilePath);

            if (seatNumbers.isEmpty()) {
                System.out.println("No seat numbers found in Excel file.");
                return;
            }

            // DEBUG: Print seat numbers
            System.out.println("\n=== Seat numbers read from Excel (" + seatNumbers.size() + ") ===");
            int index = 1;
            for (String seat : seatNumbers) {
                System.out.printf("%3d | '%s'%n", index++, seat);
            }
            System.out.println("====================================\n");

            // 2. Ensure destination folder exists
            Path destDir = Paths.get(destinationFolderPath);
            if (!Files.exists(destDir)) {
                Files.createDirectories(destDir);
                System.out.println("Created destination folder: " + destinationFolderPath);
            }

            // 3. Read all files from source folder
            File sourceFolder = new File(sourceFolderPath);
            File[] files = sourceFolder.listFiles();

            if (files == null || files.length == 0) {
                System.out.println("Source folder is empty or cannot be read: " + sourceFolderPath);
                return;
            }

            System.out.println("Total files found in source folder: " + files.length);

            int copiedCount = 0;
            int matchedFiles = 0;

            // Track matched seats
            Set<String> matchedSeatNumbers = new HashSet<>();

            // 4. Process each file
            for (File file : files) {

                if (file.isDirectory())
                    continue;

                String fileName = file.getName();
                System.out.println("Checking: " + fileName);

                boolean foundMatch = false;
                String matchedSeat = null;

                for (String seat : seatNumbers) {
                    if (fileName.contains(seat)) {
                        foundMatch = true;
                        matchedSeat = seat;
                        break;
                    }
                }

                if (foundMatch) {
                    matchedFiles++;
                    matchedSeatNumbers.add(matchedSeat);

                    Path destPath = Paths.get(destinationFolderPath, fileName);
                    Files.copy(file.toPath(), destPath, StandardCopyOption.REPLACE_EXISTING);

                    System.out.println("COPIED: " + fileName + " (matched seat: " + matchedSeat + ")");
                    copiedCount++;
                }
            }

            // ===== Summary =====
            System.out.println("\n=== Summary ===");
            System.out.println("Total seat numbers in Excel: " + seatNumbers.size());
            System.out.println("Total files checked:         " + files.length);
            System.out.println("Files that matched a seat:   " + matchedFiles);
            System.out.println("Total files copied:          " + copiedCount);

            // ===== File Not Found Seat Numbers =====
            System.out.println("\n=== File NOT Found Seat Numbers ===");

            Set<String> notFoundSeats = new HashSet<>(seatNumbers);
            notFoundSeats.removeAll(matchedSeatNumbers);

            if (notFoundSeats.isEmpty()) {
                System.out.println("All seat numbers have matching files.");
            } else {
                for (String seat : notFoundSeats) {
                    System.out.println("File not found for seat: " + seat);
                }
            }

            System.out.println("Total seats with no files: " + notFoundSeats.size());

        } catch (Exception e) {
            System.err.println("Error occurred:");
            e.printStackTrace();
        }
    }

    private static Set<String> readSeatNumbersFromExcel(String excelFilePath) throws IOException {

        Set<String> seatNumbers = new HashSet<>();

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {

                    if (cell == null)
                        continue;

                    String seat = null;

                    if (cell.getCellType() == CellType.STRING) {
                        seat = cell.getStringCellValue().trim();
                    }
                    else if (cell.getCellType() == CellType.NUMERIC) {
                        seat = String.valueOf((long) cell.getNumericCellValue());
                    }
                    else if (cell.getCellType() == CellType.FORMULA) {

                        switch (cell.getCachedFormulaResultType()) {
                            case STRING:
                                seat = cell.getStringCellValue().trim();
                                break;

                            case NUMERIC:
                                seat = String.valueOf((long) cell.getNumericCellValue());
                                break;

                            default:
                                break;
                        }
                    }

                    if (seat != null && !seat.isEmpty()) {
                        seatNumbers.add(seat);
                    }
                }
            }
        }

        return seatNumbers;
    }
}