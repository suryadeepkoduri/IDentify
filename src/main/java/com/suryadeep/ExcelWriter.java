package com.suryadeep;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.io.File;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

    String getDate() {
        LocalDate currenDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
        String formattedDate = currenDate.format(formatter);

        return formattedDate;
    }

    String getTimeStamp() {
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("HH:mm:ss dd-MM-yyyy");
        String timeStamp = now.format(formatter);

        return timeStamp;
    }

    // Method to create a new Excel Workbook by given name
    @SuppressWarnings("resource")
    void createExcelWorkbook(String name) throws FileNotFoundException, IOException {
        name = addXLSXExtension(name);

        File f = new File(name);
        if (f.exists()) {
            return;
        }

        XSSFWorkbook wb = new XSSFWorkbook();
        try (FileOutputStream fileOut = new FileOutputStream(name)) {
            wb.write(fileOut);
        }
    }

    // Method to create a Excel Worksheet in Excel Workbook
    @SuppressWarnings("unused")
    void createExcelWorksheet(String sheetName, String workbookName) {
        workbookName = addXLSXExtension(workbookName);

        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            XSSFSheet sheet1 = wb.createSheet(sheetName);
            FileOutputStream fileOut = new FileOutputStream(workbookName);
            wb.write(fileOut);
            fileOut.close();
            wb.close();
        } catch (Exception ex) {
            System.out.println("Error while creating sheet");
        }
    }

    @SuppressWarnings("unused")
    void createExcelWorksheet(String sheetName) {
        sheetName = addXLSXExtension(sheetName);

        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            XSSFSheet sheet1 = wb.createSheet(sheetName);
            FileOutputStream fileOut = new FileOutputStream(sheetName);
            wb.write(fileOut);
            fileOut.close();
            wb.close();
        } catch (Exception ex) {
            System.out.println("Error while creating sheet");
        }
    }

    String addXLSXExtension(String name) {
        if (FilenameUtils.getExtension(name) != ".xlsx") {
            name = name + ".xlsx";
            return name;
        } else {
            return name;
        }
    }

    public ExcelWriter() {
        createExcelWorksheet(getDate());
    }

}
