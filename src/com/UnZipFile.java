package com;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class UnZipFile {

	/** 
     * 解压文件到指定目录 
     * 解压后的文件名，和之前一致 
     * @param zipFile   待解压的zip文件 
     * @param descDir   指定目录 
     */  
    public void unZipFiles(String yesterday) throws IOException {
    	String zipFilePath="D:/DailyReportResouceFiles/"+yesterday+"/COSCON Network Utilization.zip";
    	File zipFile=new File(zipFilePath);
    	if(!zipFile.exists()){
    		System.out.println("zipFilePath :"+zipFilePath+"does not exits");
    	    return;
    	}
    	String descDir="D:/DailyReportResouceFiles/"+yesterday+"/";
        ZipFile zip = new ZipFile(zipFile);//指定编码格式，可以解决中文文件夹乱码  
        String name = zip.getName().substring(zip.getName().lastIndexOf('\\')+1, zip.getName().lastIndexOf('.'));  
          
        File pathFile = new File(descDir+name);  
        if (!pathFile.exists()) {  
            pathFile.mkdirs();  
        }  
          
        for (Enumeration<? extends ZipEntry> entries = zip.entries(); entries.hasMoreElements();) {  
            ZipEntry entry = (ZipEntry) entries.nextElement();  
            String zipEntryName = entry.getName();  
            InputStream in = zip.getInputStream(entry);  
            String outPath = (descDir + name +"/"+ zipEntryName).replaceAll("\\*", "/");  
              
            // 判断路径是否存在,不存在则创建文件路径  
            File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));  
            if (!file.exists()) {  
                file.mkdirs();  
            }  
            // 判断文件全路径是否为文件夹,如果是上面已经上传,不需要解压  
            if (new File(outPath).isDirectory()) {  
                continue;  
            }  
            // 输出文件路径信息  
//          System.out.println(outPath);  
  
            FileOutputStream out = new FileOutputStream(outPath);  
            byte[] buf1 = new byte[1024];  
            int len;  
            while ((len = in.read(buf1)) > 0) {  
                out.write(buf1, 0, len);  
            }  
            in.close();  
            out.close();  
        }  
        System.out.println("******************解压完毕********************");  
        return;  
    }  
      
}
