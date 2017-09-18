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
     * ��ѹ�ļ���ָ��Ŀ¼ 
     * ��ѹ����ļ�������֮ǰһ�� 
     * @param zipFile   ����ѹ��zip�ļ� 
     * @param descDir   ָ��Ŀ¼ 
     */  
    public void unZipFiles(String yesterday) throws IOException {
    	String zipFilePath="D:/DailyReportResouceFiles/"+yesterday+"/COSCON Network Utilization.zip";
    	File zipFile=new File(zipFilePath);
    	if(!zipFile.exists()){
    		System.out.println("zipFilePath :"+zipFilePath+"does not exits");
    	    return;
    	}
    	String descDir="D:/DailyReportResouceFiles/"+yesterday+"/";
        ZipFile zip = new ZipFile(zipFile);//ָ�������ʽ�����Խ�������ļ�������  
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
              
            // �ж�·���Ƿ����,�������򴴽��ļ�·��  
            File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));  
            if (!file.exists()) {  
                file.mkdirs();  
            }  
            // �ж��ļ�ȫ·���Ƿ�Ϊ�ļ���,����������Ѿ��ϴ�,����Ҫ��ѹ  
            if (new File(outPath).isDirectory()) {  
                continue;  
            }  
            // ����ļ�·����Ϣ  
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
        System.out.println("******************��ѹ���********************");  
        return;  
    }  
      
}
