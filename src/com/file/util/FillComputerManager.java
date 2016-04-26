package com.file.util;

import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellStyle;

import com.file.forms.Computer;

public class FillComputerManager {

	public static void fillReport(HSSFSheet worksheet, int startRowIndex,  
            int startColIndex, List<Computer> datasource) {  
  
        // Row offset  
        startRowIndex += 2;  
  
        // Create cell style for the body  
        HSSFCellStyle bodyCellStyle = worksheet.getWorkbook().createCellStyle();        
        bodyCellStyle.setAlignment(CellStyle.ALIGN_CENTER);  
        bodyCellStyle.setWrapText(false); //是否自动换行.  
        /**
         * List<Computer> datasource写入Excel
         */
        // Create body  
        for (int i=startRowIndex; i+startRowIndex-2< datasource.size()+2; i++) {  
            // Create a new row  
            HSSFRow row = worksheet.createRow((short) i+1);  
  
            // Retrieve the id value  
            HSSFCell cell1 = row.createCell(startColIndex+0);  
            cell1.setCellValue(datasource.get(i-2).getId());  
            cell1.setCellStyle(bodyCellStyle);  
  
            // Retrieve the brand value  
            HSSFCell cell2 = row.createCell(startColIndex+1);  
            cell2.setCellValue(datasource.get(i-2).getBrand());  
            cell2.setCellStyle(bodyCellStyle);  
  
            // Retrieve the model value  
            HSSFCell cell3 = row.createCell(startColIndex+2);  
            cell3.setCellValue(datasource.get(i-2).getCpu());  
            cell3.setCellStyle(bodyCellStyle);  
  
            // Retrieve the maximum power value  
            HSSFCell cell4 = row.createCell(startColIndex+3);  
            cell4.setCellValue(datasource.get(i-2).getGpu());  
            cell4.setCellStyle(bodyCellStyle);  
  
            // Retrieve the price value  
            HSSFCell cell5 = row.createCell(startColIndex+4);  
            cell5.setCellValue(datasource.get(i-2).getMemory());  
            cell5.setCellStyle(bodyCellStyle);  
          
            // Retrieve the efficiency value  
            HSSFCell cell6 = row.createCell(startColIndex+5);  
            cell6.setCellValue(datasource.get(i-2).getPrice());  
            cell6.setCellStyle(bodyCellStyle);  
        }  
    }  
}
