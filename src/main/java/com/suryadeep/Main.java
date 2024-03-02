package com.suryadeep;

import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {
    @SuppressWarnings("unused")
    public static void main(String[] args) throws FileNotFoundException, IOException {
        // ExcelWriter ew = new ExcelWriter();

        // System.out.println(ew.getDate());
        // //ew.createExcelWorkbook(ew.getDate());
        // //ew.createExcelWorksheet("hello", ew.getDate());
        // System.out.println(ew.getTimeStamp());
        // ew.createExcelWorkbook(ew.getDate());
        // ew.createExcelWorksheet("sheet-1", ew.getDate());
        // ew.createExcelWorksheet("sheet-2", ew.getDate());

        ExcelWriter2 ew2 = new ExcelWriter2();
        // ew2.enterData("1", "alex");
        // ew2.enterData("2", "john");
        //ew2.enterData("3", "alexa");
        ew2.enterData("4", "steve");
        ew2.enterData("5", "adam");
    }
}