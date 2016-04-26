package test;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.file.controller.PoiController;
import com.file.forms.Computer;
import com.file.service.PoiService;
import com.file.util.FileZip;

public class TestExport {
    
	
	@Test
	public void test() throws IOException {
		ArrayList<Computer> computerResult = new ArrayList<Computer>();
		for(int i=0;i<60000*2;i++){
			Computer computer = new Computer();
			computer.setId(i);
			computer.setBrand("lenovo");
			computer.setCpu("4g");
			computer.setGpu("Titan");
			computer.setMemory("1TB");
			computer.setPrice(10345.00);
			computerResult.add(computer);
			computer = null;
		}
		List<String> fileNames = new ArrayList();// 用于存放生成的文件名称
		
		File zip = new File("D://"+ "TEST" + ".zip");// 压缩文件  
		int length  = 60000;
		// 生成excel  
		for (int j = 0, n = computerResult.size() / length + 1; j < n; j++) {  
			Workbook book = new HSSFWorkbook();  
			Sheet sheet = book.createSheet("Computer");  

			double d = 0;// 用来统计  
			String file =  "D://" + "test" + "-" + j  
					+ ".xls";  

			fileNames.add(file);  
			FileOutputStream o = null;  
			try {  
				o = new FileOutputStream(file);  

				// sheet.addMergedRegion(new  
				// CellRangeAddress(list.size()+1,0,list.size()+5,6));  
				Row row = sheet.createRow(0);  
				//可以从sql中得到字段名
				row.createCell(0).setCellValue("编号");  
				row.createCell(1).setCellValue("价格");  
				row.createCell(2).setCellValue("CPU");  
				row.createCell(3).setCellValue("GPU");  
				row.createCell(4).setCellValue("品牌");  

				int m = 1;  

				for (int i = 1, min = (computerResult.size() - j * length + 1) > (length + 1) ? (length + 1)  
						: (computerResult.size() - j * length + 1); i < min; i++) {  
					m++;  
					Computer user = computerResult.get(length * (j) + i - 1);  
					Double dd = user.getPrice();  
					if (dd == null) {  
						dd = 0.0;  
					}  
					d += dd;  
					row = sheet.createRow(i);  
					row.createCell(0).setCellValue(user.getId());  
					row.createCell(1).setCellValue(user.getPrice());  
					row.createCell(2).setCellValue(user.getBrand());  
					row.createCell(3).setCellValue(user.getCpu());  
					row.createCell(4).setCellValue(dd);  

				}  
				CellStyle cellStyle2 = book.createCellStyle();  
				cellStyle2.setAlignment(CellStyle.ALIGN_CENTER);  
				row = sheet.createRow(m);  
				Cell cell0 = row.createCell(0);  
				cell0.setCellValue("Total");  
				cell0.setCellStyle(cellStyle2);  
				Cell cell4 = row.createCell(4);  
				cell4.setCellValue(d);  
				cell4.setCellStyle(cellStyle2);  
				sheet.addMergedRegion(new CellRangeAddress(m, m, 0, 3));  
			} catch (Exception e) {  
				e.printStackTrace();  
			}  
			try {  
				book.write(o);  
			} catch (Exception ex) {  
				ex.printStackTrace();  
			} finally {  
				o.flush();  
				o.close();  
			}  
		}  
		File srcfile[] = new File[fileNames.size()];  
		for (int i = 0, n = fileNames.size(); i < n; i++) {  
			srcfile[i] = new File(fileNames.get(i));  
		}  
		//压缩

		FileZip.ZipFiles(srcfile, zip);  
//		FileInputStream inStream = new FileInputStream(zip);  
//		byte[] buf = new byte[4096];  
//		int readLength;  
//		while (((readLength = inStream.read(buf)) != -1)) {  
//			
//			out.write(buf, 0, readLength);  
//		}  
//		inStream.close();  
		
	}

}
