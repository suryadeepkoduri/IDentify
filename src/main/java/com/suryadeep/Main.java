package com.suryadeep;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;

public class Main {
    @SuppressWarnings("unused")
    public static void main(String[] args) throws FileNotFoundException, IOException {
        ExcelReader er = new ExcelReader("Data.xlsx");
        Cell data = er.readCell(3, 1);
        System.err.println(data.toString());
        ExcelWriter ew = new ExcelWriter();

        
}

    void hashCalculate(Cell cell) {
        String data = cell.toString();       
        String[] arrSplit = data.split("(?<=\\d)(?=A)|(?<=A)(?=\\d)");
        int year = Integer.parseInt(arrSplit[0]); 
    }

}