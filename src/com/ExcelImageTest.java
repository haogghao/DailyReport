package com;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelImageTest {  
	public void addNetwork(String targetFile,String sheetName) throws IOException{
		FileOutputStream fileOut = null;     
        BufferedImage bufferImg = null;     
       //先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray    
  //     try {  
           ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();     
           bufferImg = ImageIO.read(new File("D:/DailyReportResouceFiles/20170912/CT2.png"));     
           ImageIO.write(bufferImg, "png", byteArrayOut);  
             
       	   FileInputStream fis = new FileInputStream(targetFile); 
           XSSFWorkbook wb = new XSSFWorkbook(fis);  
           XSSFSheet sheet1 = wb.getSheet(sheetName);   
           //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）  
           XSSFDrawing patriarch = sheet1.createDrawingPatriarch();     
           //anchor主要用于设置图片的属性  
           HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 255, 255,(short) 1, 1, (short) 5, 8);     
           anchor.setAnchorType(3);     
           //插入图片    
           patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));   
           fileOut = new FileOutputStream("D:/2.xlsx");     
           // 写入excel文件     
            wb.write(fileOut);     
            System.out.println("----Excle文件已生成------");  
//       } catch (Exception e) {  
//           e.printStackTrace();  
//       }finally{  
//           if(fileOut != null){  
//                try {  
//                   fileOut.close();  
//               } catch (IOException e) {  
//                   e.printStackTrace();  
//               }  
//           }  
//       }  
		
	}
    public static void main(String[] args) {  
    	ExcelImageTest EI=new ExcelImageTest();
    	try {
			EI.addNetwork("d:/22.xlsx", "Network");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }  
}  