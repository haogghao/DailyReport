package com;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ServerPerformance {
	public void addServerPerformance(String excelPath,String yesterday,String ytd)  throws IOException {
		BufferedReader br = null;  
		String csvPath = "D:/DailyReportResourceFiles/"+yesterday+"/CS2-ACZ-COSCON-PROD.csv";
        File f=new File(csvPath);
        if(!f.exists()){
        	System.out.println("csvPath :"+csvPath+" does not exits");
        	return;
        }
		String line ="";  
        String csvSplitBy = ",(?=([^\"]*\"[^\"]*\")*[^\"]*$)";
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
        //获取sheet
        XSSFSheet sheet = wb.getSheet("ServerPerformance");

        //创建样式
        CellStyle style = wb.createCellStyle(); 
        style.setBottomBorderColor(IndexedColors.RED.getIndex());   
        style.setLeftBorderColor(IndexedColors.RED.getIndex());
        style.setRightBorderColor(IndexedColors.RED.getIndex());    
        style.setTopBorderColor(IndexedColors.RED.getIndex());  
        style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框  
        style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框  
        style.setBorderTop(CellStyle.BORDER_THIN);// 上边框  
        style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
        
        //给单子名称一个长度
        sheet.setDefaultColumnWidth((short)40);
        sheet.setDefaultRowHeight((short) 500);
        //获取数据行数
        int rowNum=sheet.getLastRowNum();
        if(rowNum!=0){
        	rowNum+=2;
        }

        String dateStr=ytd+" 00:00:00  to "+ytd+" 23:59:59 HKT";
        XSSFRow dateRow =sheet.createRow(rowNum);
        XSSFCell dateCell=dateRow.createCell((short) 0);
        dateCell.setCellValue(dateStr);
        rowNum+=1;
        System.out.println(rowNum);
        for (int i = 0; i < dataList.size(); i++) {
            // 创建行
            XSSFRow row = sheet.createRow(rowNum+i);
            List<String> list = dataList.get(i);
            for (int j = 0; j < list.size(); j++) {
                // 创建单元格
                XSSFCell cell = row.createCell(j);
                cell.setCellStyle(style);
                cell.setCellValue(list.get(j).replace("\"", ""));
            }
        }
        // 写入到文件里面
        FileOutputStream out = new FileOutputStream(excelPath);
        wb.write(out);
        System.out.println("成功添加ServerPerformance");
        out.flush();
        out.close();
	}
}