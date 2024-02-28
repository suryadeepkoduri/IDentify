package com.suryadeep;

import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        ExcelWriter ew = new ExcelWriter();

        System.out.println(ew.getDate());
        ew.createExcelWorkbook(ew.getDate());
        ew.createExcelWorksheet("hello", ew.getDate());
    }
}