package com;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
    public static String excelPath="D:/DailyReportResouceFiles/Report/22.xlsx";
	public void clearSheet(String excelPath,String sheetName) {
        try { 
            FileInputStream fis = new FileInputStream(excelPath); 
            XSSFWorkbook wb = new XSSFWorkbook(fis); 
            XSSFSheet ws = wb.getSheet(sheetName);
            //创建样式
            XSSFCellStyle style = wb.createCellStyle(); 
            
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());   
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());    
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());  
            style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框  
            style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框  
            style.setBorderTop(CellStyle.BORDER_THIN);// 上边框  
            style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
            style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            //单元格填充黄色样式
            XSSFCellStyle dateStyle=(XSSFCellStyle) style.clone();
            dateStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            dateStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
            
            
            if(sheetName.equals("UserSync")){//清空数据及样式
            	System.out.println("UserSync");
            	
            	for(int i=4;i<=13;i++){
            		XSSFRow tempRow = ws.getRow(i);
            		for(int j=6;j<=12;j++){
            			if(j==9) continue;
            			XSSFCell xc=tempRow.createCell(j);
            			xc.setCellStyle(style);
            			xc.setCellValue("");
            		}
            	}
            	for(int i=18;i<=27;i++){
            		XSSFRow tempRow = ws.createRow(i);
            		for(int j=0;j<=12;j++){
            			if(j==3||j==4||j==5||j==9) continue;
            			XSSFCell xc=tempRow.createCell(j);
            			xc.setCellStyle(style);
            			xc.setCellValue("");
            		}
            	}
            	for(int i=32;i<=41;i++){
            		XSSFRow tempRow = ws.createRow(i);
            		for(int j=0;j<=3;j++){
            			XSSFCell xc=tempRow.createCell(j);
            			xc.setCellStyle(style);
            			xc.setCellValue("");
            		}
            	}	
            }else if(sheetName.equals("Shipments")){
            	long date = new Date().getTime();
            	
            	for(int i=1;i<=11;i++){
            		XSSFRow tempRow = ws.getRow(i);
            		if(i==1){//设置时间
            			for(int ii=-3;ii<=3;ii++){
                    		Date as = new Date(date+(ii)*24*60*60*1000); //这里可以写入参数
                            SimpleDateFormat matter = new SimpleDateFormat("yyyy/M/dd");//yyyy/MM/dd就是2017/09/12
                            String dateStr=matter.format(as);
                            XSSFCell cell = tempRow.createCell(ii+4);
                            cell.setCellStyle(dateStyle);
                            cell.setCellValue(dateStr);
                    	}
            			continue;
            		}
            		for(int j=1;j<=7;j++){
            			XSSFCell cell = tempRow.createCell(j);
            			cell.setCellStyle(style);
            			cell.setCellValue("");
            		}
            	}
            	XSSFRow tempR = ws.getRow(16);
            	for(int i=3;i<=15;i+=2){
            		XSSFCell cell=tempR.createCell(i);
            		cell.setCellStyle(dateStyle);
            		int j=(i/2)-4;
            		DateFormat df = new SimpleDateFormat("dd-MMM",Locale.ENGLISH);
            		String timeStr=df.format(new Date().getTime()+j*24*60*60*1000);
            		cell.setCellValue(timeStr);//直接设置成08-Sep即可
            	}
            	
            	
            	
            	for(int i=18;i<=26;i++){
            		XSSFRow tempRow = ws.getRow(i);
            		for(int j=3;j<=16;j++){
            			XSSFCell cell = tempRow.createCell(j);
            			cell.setCellStyle(style);
            			cell.setCellValue("");
            		}
            	}
            }else{
            wb.removeSheetAt(wb.getSheetIndex(sheetName)); 
            wb.createSheet(sheetName);
            }
            this.fileWrite(excelPath, wb); 
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
    public void fileWrite(String excelPath,XSSFWorkbook wb) throws Exception{
        FileOutputStream fileOut = new FileOutputStream(excelPath); 
        wb.write(fileOut); 
        fileOut.flush(); 
        fileOut.close(); 
    }	
   public static void main(String[] args) throws Exception {
//	  //当天日期
//       Date date = new Date();  
//       SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");  
//       String today = sdf.format(date);  
//       
//       //前一天日期
//       Date as = new Date(date.getTime()-24*60*60*1000); //这里可以写入参数
//       SimpleDateFormat matter1 = new SimpleDateFormat("yyyyMMdd");
//	     SimpleDateFormat matter2 = new SimpleDateFormat("yyyy-MM-dd");
//       String yesterday = matter1.format(as);
// 	     String ytd = matter2.format(as);
	   
        ExcelWrite ew = new ExcelWrite(); 
//        String yesterday="20170912";
//        String ytd="2017-09-13";
        
//       //解压zip文件          
//        UnZipFile zf=new UnZipFile();
//        zf.unZipFiles(yesterday);
//      //清楚Throughout,ServerPerformance and Network,Shipments
	    ew.clearSheet(excelPath, "UserSync");
//        ew.clearSheet(excelPath, "Shipments");
//        ew.clearSheet(excelPath, "Throughout");//ServerPerformance
//        ew.clearSheet(excelPath, "ServerPerformance");
//        ew.clearSheet(excelPath, "Network");
//     //插入Throughout   
//        Throughout th=new Throughout();
//        th.addThroughout(excelPath, "Throughout",yesterday);
//     //插入ServerPerformance
//        ServerPerformance sp=new ServerPerformance();
//        sp.addServerPerformance(excelPath, yesterday,ytd);
       //插入Shipments
//		Shipments sp=new Shipments();
//		sp.writeShipments(excelPath,"Shipments",5);
       //插入UserSync
	   
    }
}
