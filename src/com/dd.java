package com;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

public class dd {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
//		Date d=new Date();
//		SimpleDateFormat sf=new SimpleDateFormat("MMM d, yyy");
//		String ss=sf.format(d);
//		System.out.println(ss);

		DateFormat df = new SimpleDateFormat("dd-MMM",Locale.ENGLISH);
		String timeStr=df.format(new Date().getTime()-24*60*60*1000);
        System.out.println(timeStr);
        for(int i=3;i<=15;i+=2){
        	int j=(i/2)-4;
        	System.out.println(j);
        }
	}

}
