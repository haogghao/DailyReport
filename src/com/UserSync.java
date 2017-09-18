package com;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UserSync {
	public void writeUserSync(String excelPath,String sheetName,int dayNum){
		String UserSyncPath="D:/DailyReportResouceFiles/20170912/Report - Coscon User Profile Sync Txn Report.xlsx";
		
        try { 
            //读取UserSync
            XSSFSheet xs=getSheet(UserSyncPath,"Sheet0");
            String[][]temp = new String[xs.getLastRowNum()+1][3];
            for(int i=1;i<=xs.getLastRowNum();i++){
            	XSSFRow xr=xs.getRow(i);
            	//System.out.println(xr.getCell(0).getStringCellValue());
            		temp[i][0]=xr.getCell(0).getStringCellValue();
            		temp[i][1]=xr.getCell(1).getStringCellValue();
            		temp[i][2]=xr.getCell(2).getStringCellValue();	
            }
            System.out.println(temp[1][0]+"  "+temp[1][1]+" "+temp[1][2]);
            //写入UserSync
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
            int columnNum=0,param=0;
            if(dayNum==5){//Friday
            	for(int i=1;i<temp.length;i++){
                	for(int j=3+i;j<=13;j++){
                		XSSFRow xr=ws.getRow(j);
                		String code=xr.getCell(0).getStringCellValue();
                		String result=xr.getCell(1).getStringCellValue();
                		XSSFCell xc=null;
                		if(temp[i][0].equals(code)&&temp[i][1].equals(result)){
                			xc=xr.createCell(2);
                			xc.setCellStyle(style);
                			xc.setCellValue(temp[i][2]);
                		}else{
                			xc=xr.createCell(2);
                			xc.setCellStyle(style);
                			xc.setCellValue("0");
                		}
                	}	
                }
            }else{
            	if(dayNum==6||dayNum==7){
                	columnNum=4*dayNum-18;
                	param=3;
                }else if(dayNum==4){
                	param=31;
                }else if(dayNum<4){
                	columnNum=4*dayNum-2;
                	if(columnNum==2) columnNum=0;
                	param=17;
                }
            	
            	for(int i=1;i<temp.length;i++){//注意,表格不能为空,否则报空指针异常
                	XSSFRow xr=ws.getRow(i+param);
                	for(int j=0;j<3;j++){
                		XSSFCell xc=xr.createCell(columnNum+j);
                		xc.setCellStyle(style);
                		xc.setCellValue(temp[i][j]);
                	}
                }
            }
            
            
            this.fileWrite(excelPath, wb); 
            fis.close(); 
        } catch (Exception e) { 
            e.printStackTrace(); 
        } 
		
	}
	public XSSFSheet getSheet(String excelPath,String sheetName) throws IOException{
		FileInputStream fis = new FileInputStream(excelPath); 
        XSSFWorkbook wb = new XSSFWorkbook(fis); 
        XSSFSheet ws = wb.getSheet(sheetName);
        if(ws.getRow(0)==null){
        	System.out.println(excelPath+" 是空表");
        }
        return ws;
	}
    public void fileWrite(String excelPath,XSSFWorkbook wb) throws Exception{
        FileOutputStream fileOut = new FileOutputStream(excelPath); 
        wb.write(fileOut); 
        fileOut.flush(); 
        fileOut.close(); 
    }	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		UserSync us=new UserSync();
		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",5);
//		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",1);
//		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",2);
//		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",3);
//		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",4);
//		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",6);
//		us.writeUserSync("D:/DailyReportResouceFiles/Report/22.xlsx","UserSync",7);
	}

}
