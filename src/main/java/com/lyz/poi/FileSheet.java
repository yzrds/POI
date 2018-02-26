package com.lyz.poi;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
/**
 * 文本提取
 */
public class FileSheet {
    public static void main(String[] args) throws Exception{
        InputStream is=new FileInputStream("f:\\11.xls");
        POIFSFileSystem fs=new POIFSFileSystem(is);
        HSSFWorkbook wb=new HSSFWorkbook(fs);

        ExcelExtractor excelExtractor=new ExcelExtractor(wb);
        excelExtractor.setIncludeSheetNames(false);// 我们不需要Sheet页的名字
        System.out.println(excelExtractor.getText());
    }
}
