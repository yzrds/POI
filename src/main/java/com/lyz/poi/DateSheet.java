package com.lyz.poi;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 时间格式的单元格
 */
public class DateSheet {
    public static void main(String[] args)  throws Exception {
        Workbook workbook=new HSSFWorkbook();
        Sheet sheet=workbook.createSheet("first Sheet");
        Row row=sheet.createRow(0);
        Cell cell=row.createCell(0);
        cell.setCellValue(new Date());

        CreationHelper createHelper=workbook.getCreationHelper();
        CellStyle cellStyle=workbook.createCellStyle(); //单元格样式类
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyy-mm-dd hh:mm:ss"));
        cell=row.createCell(1); // 第二列
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        cell=row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        FileOutputStream fileOut=new FileOutputStream("F:\\workbook1.xls");
        workbook.write(fileOut);
        fileOut.close();
    }
}
