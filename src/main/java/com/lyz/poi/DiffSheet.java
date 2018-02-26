package com.lyz.poi;

import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 不同内容格式的单元格
 */
public class DiffSheet {
    public static void main(String[] args) throws Exception{
        Workbook wb=new HSSFWorkbook();
        Sheet sheet=wb.createSheet(" fiest Sheet");
        Row row=sheet.createRow(0); // 创建一个行
        Cell cell=row.createCell(0); // 创建一个单元格  第1列
        cell.setCellValue(new Date().toString());

        row.createCell(1).setCellValue(1);
        row.createCell(2).setCellValue("str");
        row.createCell(3).setCellValue(true);
        row.createCell(4).setCellValue(HSSFCell.CELL_TYPE_NUMERIC);
        row.createCell(5).setCellValue(false);

        FileOutputStream fileOut=new FileOutputStream("F:\\workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
