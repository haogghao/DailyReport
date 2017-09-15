package com;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
public class Datetime {

	/**
	 * @param args
	 */
	public static void main(String[] args)throws ParseException  {
		// TODO Auto-generated method stub
		//当天日期
        Date date = new Date();  
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");  
        String today = sdf.format(date);  
        System.out.println("格式化后的日期：" + today);
        
        //前一天日期
        Date as = new Date(date.getTime()-24*60*60*1000); //这里可以写入参数
        SimpleDateFormat matter1 = new SimpleDateFormat("yyyyMMdd");
        String yesterday = matter1.format(as);
        System.out.println(yesterday);
	}

}
