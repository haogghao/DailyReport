package com;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public void clearSheet(String targetFile,String sheetName) { 
        try { 
            FileInputStream fis = new FileInputStream(targetFile); 
            XSSFWorkbook wb = new XSSFWorkbook(fis); 
            XSSFSheet ws = wb.getSheet(sheetName);
            wb.removeSheetAt(wb.getSheetIndex(sheetName)); 
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
		
   public static void main(String[] args) throws IOException {
	  //当天日期
       Date date = new Date();  
       SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");  
       String today = sdf.format(date);  
       System.out.println("今天:" + today);
       
       //前一天日期
       Date as = new Date(date.getTime()-24*60*60*1000); //这里可以写入参数
       SimpleDateFormat matter1 = new SimpleDateFormat("yyyyMMdd");
       String yesterday = matter1.format(as);
       System.out.println("昨天:"+yesterday);
	   
        ExcelWrite ew = new ExcelWrite(); 
        String excelPath="D:/DailyReportResouceFiles/Report/22.xlsx";
        String csvPath = "D:/DailyReportResouceFiles/20170912/CS2-ACZ-COSCON-PROD.csv";  
     //清楚Throughout,ServerPerformance and Network
//        ew.clearSheet(excelPath, "Throughout");//ServerPerformance
//        ew.clearSheet(excelPath, "ServerPerformance");
//        ew.clearSheet(excelPath, "Network");
     //插入Throughout   
//        Throughout th=new Throughout();
//        th.addThroughout(excelPath, "Throughout");
     //插入ServerPerformance
//        ServerPerformance sp=new ServerPerformance();
//        sp.addServerPerformance(excelPath, csvPath);
        
        
        
    }
}
