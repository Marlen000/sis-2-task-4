package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ScholarshipProcessor {

    public static void main(String[] args) {
        String inputFilePath = "C:/Users/marle/IdeaProjects/StudentScholarshipProcessor/src/studens.xlsx";
        String outputFilePath = "C:/Users/marle/IdeaProjects/StudentScholarshipProcessor/src/updated_studens.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            Sheet updatedSheet = workbook.createSheet("Updated Students");

            Row headerRow = updatedSheet.createRow(0);
            Row originalHeaderRow = sheet.getRow(0);
            for (int i = 0; i < originalHeaderRow.getPhysicalNumberOfCells(); i++) {
                headerRow.createCell(i).setCellValue(originalHeaderRow.getCell(i).getStringCellValue());
            }
            headerRow.createCell(originalHeaderRow.getPhysicalNumberOfCells()).setCellValue("Scholarship");

            boolean isFirstRow = true;
            int rowNum = 1;
            for (Row row : sheet) {
                if (isFirstRow) {
                    isFirstRow = false;
                    continue;
                }

                Cell lastNameCell = row.getCell(0);
                Cell firstNameCell = row.getCell(1);
                Cell gradeCell = row.getCell(2);

                if (lastNameCell == null || firstNameCell == null || gradeCell == null) {
                    continue;
                }

                String lastName = lastNameCell.getStringCellValue();
                String firstName = firstNameCell.getStringCellValue();
                double grade = gradeCell.getNumericCellValue();

                String scholarship = grade >= 70 ? "Stepa bar)" : "Stepa zhok(";

                Row updatedRow = updatedSheet.createRow(rowNum++);
                updatedRow.createCell(0).setCellValue(lastName);
                updatedRow.createCell(1).setCellValue(firstName);
                updatedRow.createCell(2).setCellValue(grade);
                updatedRow.createCell(3).setCellValue(scholarship);
            }
            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                workbook.write(fos);
            }

            workbook.close();
            System.out.println("Новый файл с обновленными данными был создан: " + outputFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
