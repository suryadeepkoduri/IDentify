package com.suryadeep;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter2 {

    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet("Sheet-1");
    
    ExcelWriter2() throws FileNotFoundException, IOException {
        String name = addXLSXExtension(getDate());

        File f = new File(name);
        if (f.exists()) {
            return ;
        }

        XSSFWorkbook wb = new XSSFWorkbook();
        try (FileOutputStream fileOut = new FileOutputStream(name)) {
            wb.write(fileOut);
            System.out.println("File Created by constructor");
        }
    }





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


    String addXLSXExtension(String name) {
        if (FilenameUtils.getExtension(name) != ".xlsx") {
            name = name + ".xlsx";
            return name;
        } else {
            return name;
        }
    }

}
