package com.suryadeep;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

    private Workbook workbook = new XSSFWorkbook();
    private Sheet sheet = workbook.createSheet("Sheet-1");
    private int rowNo = 0;

    // Corrected by chatgpt
    /*
     * Notes
     * 
     *  ExcelWriter2() throws FileNotFoundException, IOException {
            String name = addXLSXExtension(getDate());
            sheetHeader();

            File f = new File(name);
            if (f.exists()) {
                return;
            }


            try (FileOutputStream fileOut = new FileOutputStream(name)) {
                workbook.write(fileOut);
                System.out.println("File Created by constructor");
            }
        }

     * 
     * The constructor is creating a new workbook and sheet every time it's called, even if the file already exists. 
     * It should only create a new workbook if the file does not exist. Additionally, it should load an existing workbook if the file already exists.
     * 
     * 
     */
    public ExcelWriter() throws IOException {
        String name = addXLSXExtension(DateTime.getDate());

        File f = new File(name);
        if (!f.exists()) {
            sheetHeader();

            try (FileOutputStream fileOut = new FileOutputStream(name)) {
                workbook.write(fileOut);
                System.out.println("File Created by constructor");
            }
        } else {
            try (FileInputStream fileIn = new FileInputStream(name)) {
                workbook = new XSSFWorkbook(fileIn);
                sheet = workbook.getSheetAt(0);
            }
        }
    }


    public void enterData(String rollNo, String name) throws IOException {
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        row.createCell(0).setCellValue(rollNo);
        row.createCell(1).setCellValue(name);
        row.createCell(2).setCellValue(DateTime.getTimeStamp());

        try (FileOutputStream fileOut = new FileOutputStream(addXLSXExtension(DateTime.getDate()))) {
            workbook.write(fileOut);
        }
    }


    private void sheetHeader() {
        Row row = sheet.createRow(rowNo++);
        row.createCell(0).setCellValue("Roll No");
        row.createCell(1).setCellValue("Name");
        row.createCell(2).setCellValue("Time Stamp");
    }

    
    private String addXLSXExtension(String name) {
        if (!FilenameUtils.getExtension(name).equalsIgnoreCase("xlsx")) {
            name = name + ".xlsx";
        }
        return name;
    }

}

class DateTime {

    public static String getTimeStamp() {
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("HH:mm:ss dd-MM-yyyy");
        String timeStamp = now.format(formatter);

        return timeStamp;
    }

    public static String getDate() {
        LocalDate currenDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
        String formattedDate = currenDate.format(formatter);

        return formattedDate;
    }

}
