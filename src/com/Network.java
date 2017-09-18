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

public class Network {
	public void addNetwork(String excelPath,String picturePath,String sheetName,int days) throws IOException{
		FileOutputStream fileOut = null;     
        BufferedImage min5 = null; 
        BufferedImage min30 = null; 
        File f=new File(picturePath);
        if(!f.exists()){
        	System.out.println("the network picture path:"+picturePath+" does not exits");
        	return;
        }
        String pic5minPath=picturePath+"/COSCON Network Utilization/5min.png";
        String pic30minPath=picturePath+"/COSCON Network Utilization/30min.png";
        String dateTile="2017/Sep/08 COSCON 10M lease line usage : < 25%";
        String title5="Daily (5 minutes average)";
        String title30="Weekly (30 minutes average)";
        int row1,column1,row2,column2,row3,column3,row4,column4;
        if(days>4){
        	row1=2;column1=(days-5)*10;
        	row2=13;column2=(days-5)*10+8;
        }else if(days<4){
        	row1=33;column1=(days-1)*10;
        	row2=45;column2=(days-1)*10+8;
        }else{
        	row1=64;column1=0;row2=76;column2=8;
        }
    	row3=row1+15;column3=column1;
    	row4=row2+15;column4=column2;
       //先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray    
       try {  
           ByteArrayOutputStream byteArrayOut5 = new ByteArrayOutputStream(); 
           ByteArrayOutputStream byteArrayOut30 = new ByteArrayOutputStream();
           min5 = ImageIO.read(new File(pic5minPath)); 
           min30 = ImageIO.read(new File(pic30minPath)); 
           ImageIO.write(min5, "png", byteArrayOut5);  //指定图片格式
           ImageIO.write(min30, "png", byteArrayOut30);
       	   FileInputStream fis = new FileInputStream(excelPath); 
           XSSFWorkbook wb = new XSSFWorkbook(fis);  
           XSSFSheet sheet1 = wb.getSheet(sheetName);   
           //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）  
           XSSFDrawing patriarch = sheet1.createDrawingPatriarch();     
           //anchor主要用于设置图片的属性  0 0,255 255应该表示图片插入的范围;0 1,25 18 表示图片开始和结束位置
           XSSFClientAnchor anchor5 = new XSSFClientAnchor(0, 0, 255, 255,column1, row1, column2, row2);   //5,8  
           XSSFClientAnchor anchor30 = new XSSFClientAnchor(0, 0, 255, 255,column3, row3, column4, row4);   //5,8  

           //插入图片    
           patriarch.createPicture(anchor5, wb.addPicture(byteArrayOut5.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG));
           patriarch.createPicture(anchor30, wb.addPicture(byteArrayOut30.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG)); 
           fileOut = new FileOutputStream(excelPath);     //实际保存的地方,不要写错
           // 写入excel文件     
            wb.write(fileOut);     
            System.out.println("----Network图片插入成功------");  
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
