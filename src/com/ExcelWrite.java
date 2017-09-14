package com;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
    /** 
     * 删除指定的Sheet 
     * @param targetFile  目标文件 
     * @param sheetName   Sheet名称 
     */ 
    public void deleteSheet(String targetFile,String sheetName) { 
        try { 
            FileInputStream fis = new FileInputStream(targetFile); 
            XSSFWorkbook wb = new XSSFWorkbook(fis); 
            //wb.createSheet("gungun");
            XSSFSheet ws = wb.getSheetAt(1);
            //XSSFSheet ws = wb.getSheet(sheetName);
            //wb.removeSheetAt(wb.getSheetIndex(sheetName)); 
            this.fileWrite(targetFile, wb); 
            fis.close(); 
        } catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }  
    public void createSheet(String targetFile,String sheetName) { 
        try { 
        	FileInputStream fis = new FileInputStream(targetFile); 
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            wb.createSheet(sheetName);
            this.fileWrite(targetFile, wb); 
            fis.close();
        } catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }
    /** 
     * 写删除后的Excel文件 
     * @param targetFile  目标文件 
     * @param wb          Excel对象 
     * @throws Exception 
     */ 
    public void fileWrite(String targetFile,XSSFWorkbook wb) throws Exception{
        FileOutputStream fileOut = new FileOutputStream(targetFile); 
        wb.write(fileOut); 
        fileOut.flush(); 
        fileOut.close(); 
    }
   public static void main(String[] args) { 
        ExcelWrite ew = new ExcelWrite(); 
        ew.createSheet("d:/22.xlsx", "baba");
        ew.deleteSheet("d:/22.xlsx", "Network");
        System.out.println("成功");
    }
}
