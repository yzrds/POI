package com.lyz.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class Cell {
    public static void main(String[] args) {
        // create a new workbook
        Workbook workbook = new HSSFWorkbook();
        workbook.createSheet("first sheet");
        workbook.createSheet("second sheet");
    }
}
