package Duplicate_pdf_file_remove;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.HashSet;
import java.util.Set;

public class RemoveDuplicateSeatPDF {

    public static void main(String[] args) throws IOException {

        String sourceDir = "F:\\GCCTBC-APR 2026\\Photocopy folder\\marathi\\New folder (2)";
        String targetDir = "F:\\GCCTBC-APR 2026\\Photocopy folder\\marathi\\PDF_UNIQUE";

        File sourceFolder = new File(sourceDir);
        File targetFolder = new File(targetDir);

        if (!targetFolder.exists()) {
            targetFolder.mkdirs();
        }

        Set<String> seatNumberSet = new HashSet<>();

        File[] files = sourceFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));

        if (files == null) {
            System.out.println("No PDF files found.");
            return;
        }

        for (File file : files) {
            String fileName = file.getName();
            String[] parts = fileName.split("_");

            if (parts.length < 2) {
                continue;
            }

            String seatNumber = parts[1];

            if (!seatNumberSet.contains(seatNumber)) {

                seatNumberSet.add(seatNumber);

                File destFile = new File(targetFolder, fileName);
                Files.copy(file.toPath(), destFile.toPath(),
                        StandardCopyOption.REPLACE_EXISTING);

                System.out.println("Copied: " + fileName);
            } else {
                System.out.println("Duplicate ignored: " + fileName);
            }
        }

        System.out.println("✅ Process completed.");
    }
}
