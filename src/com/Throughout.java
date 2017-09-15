package com;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Throughout {  
	public void addThroughout(String excelPath,String sheetName) throws IOException{
		FileOutputStream fileOut = null;     
        BufferedImage bufferImg = null;     
       //先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray    
       try {  
           ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();     
           bufferImg = ImageIO.read(new File("D:/DailyReportResouceFiles/20170912/CT2.png"));     
           ImageIO.write(bufferImg, "png", byteArrayOut);  //指定图片格式
             
       	   FileInputStream fis = new FileInputStream(excelPath); 
           XSSFWorkbook wb = new XSSFWorkbook(fis);  
           XSSFSheet sheet1 = wb.getSheet(sheetName);   
           //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）  
           XSSFDrawing patriarch = sheet1.createDrawingPatriarch();     
           //anchor主要用于设置图片的属性  0 0,255 255应该表示图片插入的范围;0 1,25 18 表示图片开始和结束位置
           XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 255, 255,(short) 0, 1, (short) 25, 18);   //5,8  
           anchor.setAnchorType(3);
           XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 255, 255,(short) 0, 20, (short) 25, 37);   //5,8  
           anchor1.setAnchorType(3); 
           XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, 255, 255,(short) 0, 40, (short) 25, 57);   //5,8  
           anchor2.setAnchorType(3); 
           //插入图片    
           patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG));
           patriarch.createPicture(anchor1, wb.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG)); 
           patriarch.createPicture(anchor2, wb.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG));
           fileOut = new FileOutputStream(excelPath);     //实际保存的地方,不要写错
           // 写入excel文件     
            wb.write(fileOut);     
            System.out.println("----BC,BL,CT图片插入成功------");  
       } catch (Exception e) {  
           e.printStackTrace();  
       }finally{  
           if(fileOut != null){  
                try {  
                   fileOut.close();  
               } catch (IOException e) {  
                   e.printStackTrace();  
               }  
           }  
       }  
		
	}
}  