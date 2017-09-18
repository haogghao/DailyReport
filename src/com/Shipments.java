package com;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Shipments {
	
	public XSSFSheet getSheet(String excelPath,String sheetName) throws IOException{
		FileInputStream fis = new FileInputStream(excelPath); 
        XSSFWorkbook wb = new XSSFWorkbook(fis); 
        XSSFSheet ws = wb.getSheet(sheetName);
        if(ws.getRow(0)==null){
        	System.out.println(excelPath+" 是空表");
        }
        return ws;
	}

	public void writeShipments(String excelPath,String sheetName,int dayNum) throws Exception { 
		String SFPath="D:/DailyReportResouceFiles/20170912/ACZone Shipment Folder Txn Report.xlsx";
		String ACZonePath="D:/DailyReportResouceFiles/20170912/ACZone TXN Monitor.xlsx";
		String UserSyncPath="D:/DailyReportResouceFiles/20170912/Report - Coscon User Profile Sync Txn Report.xlsx";
		String STDZonePath="D:/DailyReportResouceFiles/20170912/STDZone COSCON BR SI Daily TXN Report.xlsx";
		
        int []data=new int[10];
		//获取SF
            XSSFSheet SF=getSheet(SFPath,"Sheet0");
            for(int i=0;i<=SF.getLastRowNum();i++){
            	XSSFRow cell =SF.getRow(i);
            		if("2017-09-12".equals(cell.getCell(0).getStringCellValue())){
            			data[9]=Integer.parseInt(cell.getCell(1).getStringCellValue());
            			break;
            		}else{
            			if(i==SF.getLastRowNum())
            			System.out.println("SF date is null");
            		}
            }
        
		//获取STDZone
        XSSFSheet STDZone=getSheet(STDZonePath,"Sheet0");
        for(int i=0;i<=STDZone.getLastRowNum();i++){
        	XSSFRow cell =STDZone.getRow(i);
        		if("2017-09-12".equals(cell.getCell(0).getStringCellValue())){
        			if("SI".equals(cell.getCell(2).getStringCellValue())){
        				data[7]=Integer.parseInt(cell.getCell(1).getStringCellValue());
            			continue;
        			}else if("BR".equals(cell.getCell(2).getStringCellValue())){
        				data[5]=Integer.parseInt(cell.getCell(1).getStringCellValue());
            			continue;
        			}else continue;
        		}else{
        			if(i==STDZone.getLastRowNum())
        			System.out.println("STDZone date is null");
        		}
        }
    	//获取ACZone
        XSSFSheet ACZone=getSheet(ACZonePath,"Sheet0");
        for(int i=0;i<=ACZone.getLastRowNum();i++){
        	XSSFRow cell =ACZone.getRow(i);
        		if("2017-09-12".equals(cell.getCell(0).getStringCellValue())){
        			if("SI".equals(cell.getCell(2).getStringCellValue())){
        				data[6]=Integer.parseInt(cell.getCell(1).getStringCellValue());
        				data[3]=Integer.parseInt(cell.getCell(3).getStringCellValue());
        				data[2]=Integer.parseInt(cell.getCell(4).getStringCellValue());
            			continue;
        			}else if("BR".equals(cell.getCell(2).getStringCellValue())){
        				data[8]=Integer.parseInt(cell.getCell(1).getStringCellValue());
        				data[0]=Integer.parseInt(cell.getCell(3).getStringCellValue());
        				data[1]=Integer.parseInt(cell.getCell(4).getStringCellValue());
            			continue;
        			}else continue;
        			
        		}else{
        			if(i==ACZone.getLastRowNum())
        			System.out.println("ACZone date is null");
        		}
        }  
        data[1]=data[1]+data[5];//Online BR
        data[2]=data[2]+data[7];//Online SI
        //得到total的值
        for(int i=0;i<4;i++){
        	data[4]+=data[i];
        }
        
      //把数据写入Excel
    	FileInputStream fis = new FileInputStream(excelPath); 
        XSSFWorkbook wb = new XSSFWorkbook(fis); 
        XSSFSheet targetWs = wb.getSheet(sheetName);
        int column1=0,column2=0;
        if(dayNum>4){
        	column1=dayNum-4;
        }else if(dayNum<=4){
        	column1=dayNum+3;
        }else{
        	System.out.println("dayNum error");
        }
        
        column2=2*column1+1;
        for(int i=7;i<=11;i++){
        	XSSFRow xf=targetWs.getRow(i);
        	xf.getCell(column1).setCellValue(data[i-7]);
        }
        for(int i=18;i<=22;i++){
        	XSSFRow xf=targetWs.getRow(i);
        	xf.getCell(column2).setCellValue(data[i-13]);
        }   
        
        //保存修改,关闭
            this.fileWrite(excelPath, wb); 
            fis.close(); 
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

}
