package com.file.util;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFSheet;

public class Writer {

	public static void write(HttpServletResponse response, HSSFSheet worksheet) {  
		  
       
        try {  
            // Retrieve the output stream  
            ServletOutputStream outputStream = response.getOutputStream();  
            // Write to the output stream  
            worksheet.getWorkbook().write(outputStream);  
            // 清除缓存  
            outputStream.flush();  
  
        } catch (Exception e) {  
            e.printStackTrace();
        }  
    }  
}
