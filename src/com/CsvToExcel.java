package com;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CsvToExcel {
    public static void main(String[] args) throws IOException {
        // 读取csv文件
    	String csvPath = "D:/DailyReportResouceFiles/20170907/CS2-ACZ-COSCON-PROD.csv";  
    	String excelPath="D://21.xlsx";
        BufferedReader br = null;  
        String line ="";  
        String csvSplitBy = ",(?=([^\"]*\"[^\"]*\")*[^\"]*$)";
        

        PrintWriter writer_path = null;  
        List<List<String>>  dataList= new ArrayList<List<String>>();
        try {  
            br = new BufferedReader(new FileReader(csvPath));  
            while((line = br.readLine()) != null){  
                //use comma as separatpr  
                String[] major = line.split(csvSplitBy);   
                List<String> rowData = new ArrayList<String>();
                for (int i = 0; i < major.length; i++) {
                	rowData.add(major[i]);  	
                }
                dataList.add(rowData);
            }  
              
        } catch (FileNotFoundException e) {  
          
            e.printStackTrace();  
        } catch (UnsupportedEncodingException e) {  
              
            e.printStackTrace();  
        } catch (IOException e) {  
            // TODO Auto-generated catch block  
            e.printStackTrace();  
        }  
        
        // 使用poi导出excel,poi是通过循环的方式创建行和单元格
        // 声明一个工作薄
        FileInputStream fis = new FileInputStream(excelPath); 
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        //声明一个单子并命名
        XSSFSheet sheet = wb.getSheet("UserSync");
        //给单子名称一个长度
        sheet.setDefaultColumnWidth((short)15);
        for (int i = 0; i < dataList.size(); i++) {
            // 创建行
            XSSFRow row = sheet.createRow(i);
            List<String> list = dataList.get(i);
            for (int j = 0; j < list.size(); j++) {
                // 创建单元格
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(list.get(j).replace("\"", ""));
            }
        }
        
        // 写入到文件里面
        FileOutputStream out = new FileOutputStream(excelPath);
        wb.write(out);
        out.flush();
        out.close();
    }
}