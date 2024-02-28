package com.suryadeep;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.LocalDate;

public class Main {
    public static void main(String[] args) {
        ExcelWriter ew = new ExcelWriter();

        System.out.println(ew.getDate());
        try {
            ew.createExcelWorkbook("surya");
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}