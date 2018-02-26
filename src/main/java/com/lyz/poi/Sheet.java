package com.lyz.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;

public class Sheet {
    public static void main(String[] args) {
        // create a new workbook
        Workbook workbook = new HSSFWorkbook();
        workbook.createSheet("first sheet");
        workbook.createSheet("second sheet");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream("F:\\workbook1.xls");
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
