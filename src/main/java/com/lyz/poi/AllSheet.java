package com.lyz.poi;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * 遍历工作簿的行和列并获取单元格内容，如果不能运行就是excel的版本问题，wps是可以的
 */
public class AllSheet {
    public static void main(String[] args) throws Exception{
        InputStream inputStream=new FileInputStream("f:\\11.xls");
        POIFSFileSystem poifsFileSystem=new POIFSFileSystem(inputStream);
        HSSFWorkbook hssfWorkbook=new HSSFWorkbook(poifsFileSystem);
        HSSFSheet hssfSheet=hssfWorkbook.getSheetAt(0); // 获取第一个Sheet页
        if(hssfSheet==null){
            return;
        }
        // 遍历行Row
        for(int rowNum=0;rowNum<=hssfSheet.getLastRowNum();rowNum++){
            HSSFRow hssfRow=hssfSheet.getRow(rowNum);
            if(hssfRow==null){
                continue;
            }
            // 遍历列Cell
            for(int cellNum=0;cellNum<=hssfRow.getLastCellNum();cellNum++){
                HSSFCell hssfCell=hssfRow.getCell(cellNum);
                if(hssfCell==null){
                    continue;
                }
                System.out.print(" "+getValue(hssfCell));
            }
            System.out.println();
        }
    }

    private static String getValue(HSSFCell hssfCell){
        if(hssfCell.getCellType()==HSSFCell.CELL_TYPE_BOOLEAN){
            return String.valueOf(hssfCell.getBooleanCellValue());
        }else if(hssfCell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
            return String.valueOf(hssfCell.getNumericCellValue());
        }else{
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }
}
