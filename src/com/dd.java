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
		Calendar cal = Calendar.getInstance(); 
		int whichDay1 = cal.get(Calendar.DAY_OF_WEEK);
	 System.out.println(whichDay1);
	}

}
