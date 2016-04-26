package com.file.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {

	/** 
     * 获得yyyy-MM-dd格式的当前日期 
     */  
    public static String getNowTime() {  
        Date date = new Date();  
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");  
        return dateFormat.format(date);  
    }  
}
