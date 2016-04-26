package com.file.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;




public class Layouter {

	/** 
     * 创建报表 
     */  
    public static void buildReport(HSSFSheet worksheet, int startRowIndex,  
            int startColIndex) {  
        // 设置列的宽度  
        worksheet.setColumnWidth(0, 5000);  
        worksheet.setColumnWidth(1, 5000);  
        worksheet.setColumnWidth(2, 5000);  
        worksheet.setColumnWidth(3, 5000);  
        worksheet.setColumnWidth(4, 5000);  
        worksheet.setColumnWidth(5, 5000);  
  
        buildTitle(worksheet, startRowIndex, startColIndex);  
  
        buildHeaders(worksheet, startRowIndex, startColIndex);  
  
    }  
  
    /** 
     * 创建报表标题和日期 
     */  
    private static void buildTitle(HSSFSheet worksheet, int startRowIndex,  
            int startColIndex) {  
        // 设置报表标题字体  
        Font fontTitle = worksheet.getWorkbook().createFont();  
        fontTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);  
        fontTitle.setFontHeight((short) 280);  
  
        // 标题单元格样式  
        HSSFCellStyle cellStyleTitle = worksheet.getWorkbook().createCellStyle();  
        cellStyleTitle.setAlignment(CellStyle.ALIGN_CENTER);  
        cellStyleTitle.setWrapText(true);  
        cellStyleTitle.setFont((HSSFFont) fontTitle);  
  
        // 报表标题  
        HSSFRow rowTitle = worksheet.createRow((short) startRowIndex);  
        rowTitle.setHeight((short) 500);  
        HSSFCell cellTitle = rowTitle.createCell(startColIndex);  
        cellTitle.setCellValue("Computer Report!");  
        cellTitle.setCellStyle(cellStyleTitle);  
  
        // 合并区域内的报告标题  
        worksheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));  
  
        // date header  
        HSSFRow dateTitle = worksheet.createRow((short) startRowIndex + 1);  
        HSSFCell cellDate = dateTitle.createCell(startColIndex);  
        cellDate.setCellValue("这个报表创建于: " + DateUtils.getNowTime());  
    }  
  
    /** 
     * 创建表头 
     */  
    private static void buildHeaders(HSSFSheet worksheet, int startRowIndex,  
            int startColIndex) {  
        // Header字体  
       Font font = worksheet.getWorkbook().createFont();
    	
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);  
  
        // 单元格样式  
        HSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();  
        headerCellStyle.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);  
        headerCellStyle.setFillPattern(CellStyle.FINE_DOTS);  
        headerCellStyle.setAlignment(CellStyle.ALIGN_CENTER);  
        headerCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);  
        headerCellStyle.setWrapText(true);  
        headerCellStyle.setFont((HSSFFont) font);  
        headerCellStyle.setBorderBottom(CellStyle.BORDER_THIN);  
  
        // 创建字段标题  
        HSSFRow rowHeader = worksheet.createRow((short) startRowIndex + 2);  
        rowHeader.setHeight((short) 500);  
  
        HSSFCell cell1 = rowHeader.createCell(startColIndex + 0);  
        cell1.setCellValue("Id");  
        cell1.setCellStyle(headerCellStyle);  
  
        HSSFCell cell2 = rowHeader.createCell(startColIndex + 1);  
        cell2.setCellValue("Brand");  
        cell2.setCellStyle(headerCellStyle);  
  
        HSSFCell cell3 = rowHeader.createCell(startColIndex + 2);  
        cell3.setCellValue("CPU");  
        cell3.setCellStyle(headerCellStyle);  
  
        HSSFCell cell4 = rowHeader.createCell(startColIndex + 3);  
        cell4.setCellValue("GPU");  
        cell4.setCellStyle(headerCellStyle);  
  
        HSSFCell cell5 = rowHeader.createCell(startColIndex + 4);  
        cell5.setCellValue("Memory");  
        cell5.setCellStyle(headerCellStyle);  
  
        HSSFCell cell6 = rowHeader.createCell(startColIndex + 5);  
        cell6.setCellValue("Price");  
        cell6.setCellStyle(headerCellStyle);  
  
    }  
}
