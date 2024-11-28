package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ScholarshipProcessor {
    public static void main(String[] args) {
        String filePath = "C:/Users/marle/IdeaProjects/StudentScholarshipProcessor/src/studens.xlsx"; 

        try (FileInputStream fis = new FileInputStream(new File(filePath))) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            boolean isFirstRow = true;

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

                if (grade >= 70) {
                    System.out.println(firstName + " " + lastName + " Stepa bar)");
                } else {
                    System.out.println(firstName + " " + lastName + " Stepa zhok(");
                }
            }
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
