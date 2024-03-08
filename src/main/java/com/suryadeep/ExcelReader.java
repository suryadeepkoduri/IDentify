package com.suryadeep;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    private Workbook workbook = new XSSFWorkbook();
    private Sheet sheet ;

    public ExcelReader(String file) throws IOException {
        try(FileInputStream fileIn = new FileInputStream(file)) {
            workbook = new XSSFWorkbook(fileIn);
            sheet = workbook.getSheetAt(0);
        }
    }

    public Cell readCell(int cellNo, int rowNo) {
        Row row = sheet.getRow(rowNo);
        return row.getCell(cellNo);
    }
}
