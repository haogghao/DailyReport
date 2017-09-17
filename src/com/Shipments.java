package com;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Shipments {

	public void clearSheet(String excelPath,String sheetName) { 
        try { 
            FileInputStream fis = new FileInputStream(excelPath); 
            XSSFWorkbook wb = new XSSFWorkbook(fis); 
            XSSFSheet ws = wb.getSheet(sheetName);
            
            wb.removeSheetAt(wb.getSheetIndex(sheetName)); 
            wb.createSheet(sheetName);
            this.fileWrite(excelPath, wb); 
            fis.close(); 
        } catch (Exception e) { 
            e.printStackTrace(); 
        } 
    }  
    /** 
     * дɾ�����Excel�ļ� 
     * @param targetFile  Ŀ���ļ� 
     * @param wb          Excel���� 
     * @throws Exception 
     */ 
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

	}

}
