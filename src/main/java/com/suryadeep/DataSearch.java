package com.suryadeep;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class DataSearch {
    DataSearch() throws FileNotFoundException, IOException {
        try(FileInputStream fileIn = new FileInputStream("Data.xlsx")) {
            
        }
    }

    String getStudentName(String rollNo) {
        
    }
}
