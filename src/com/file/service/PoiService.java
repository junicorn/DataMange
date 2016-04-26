package com.file.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.file.dao.PoiDao;
import com.file.forms.Computer;
import com.file.util.FileZip;
import com.file.util.FillComputerManager;
import com.file.util.Layouter;
import com.file.util.Writer;

@Service("poiService")  
@Transactional  
public class PoiService {

	@Resource(name = "poiDao")
	private PoiDao dao;  
	/**
	 * excel 2007一个sheet 1048576行，16384列
	 * excel 2003 一个sheet 65536行，256列
	 * 查出的数据多了得分批导出
	 * @param response
	 */
	public void exportXLS(HttpServletResponse response) {  

		// 1.创建一个 workbook  
		HSSFWorkbook workbook = new HSSFWorkbook();  

		// 2.创建一个 worksheet  
		HSSFSheet worksheet = workbook.createSheet("Computer");  

		// 3.定义起始行和列  
		int startRowIndex = 0;  
		int startColIndex = 0;  

		// 4.创建title,data,headers  
		Layouter.buildReport(worksheet, startRowIndex, startColIndex);  

		// 5.填充数据  
		FillComputerManager.fillReport(worksheet, startRowIndex, startColIndex,  
				getDatasource());  
		int filename =(int)(1+Math.random()*(10000-1+1));
		// 6.设置reponse参数  
		String fileName = String.valueOf(filename)+".xls";  //文件名
		response.setHeader("Content-Disposition", "inline; filename="  
				+ fileName);  
		// 确保发送的当前文本格式  
		response.setContentType("application/vnd.ms-excel");  
		// 7. 输出流  
		Writer.write(response, worksheet);  

	}  

	/** 
	 * 读取报表 
	 */  
	public List<Computer> readReport(InputStream inp) {  

		List<Computer> computerList = new ArrayList<Computer>();  

		try {  
			String cellStr = null;  

			Workbook wb = WorkbookFactory.create(inp);  

			Sheet sheet = wb.getSheetAt(0);// 取得第一个sheets  

			//从第四行开始读取数据  
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {  

				Computer computer = new Computer();  
				Computer addComputer = new Computer();  

				Row row = sheet.getRow(i); // 获取行(row)对象  
				System.out.println(row);
				if (row == null) {  
					// row为空的话,不处理  
					continue;  
				}  

				for (int j = 0; j < row.getLastCellNum(); j++) {  

					Cell cell = row.getCell(j); // 获得单元格(cell)对象  

					// 转换接收的单元格  
					cellStr = ConvertCellStr(cell, cellStr);  

					// 将单元格的数据添加至一个对象  
					addComputer = addingComputer(j, computer, cellStr);  

				}  
				// 将添加数据后的对象填充至list中  
				computerList.add(addComputer);  
			}  

		} catch (InvalidFormatException e) {  
			e.printStackTrace();  
		} catch (IOException e) {  
			e.printStackTrace();  
		} finally {  
			if (inp != null) {  
				try {  
					inp.close();  
				} catch (IOException e) {  
					e.printStackTrace();  
				}  
			} else {  

			}  
		}  
		return computerList;  

	}  

	/** 
	 * 从数据库获得所有的Computer信息. 
	 */  
	private List<Computer> getDatasource() {  
		return dao.getComputer();  
	}  

	/** 
	 * 读取报表的数据后批量插入 
	 */  
	public int[] insertComputer(List<Computer> list) {  
		return dao.insertComputer(list);  

	}  

	/** 
	 * 获得单元格的数据添加至computer 
	 *  
	 * @param j 
	 *            列数 
	 * @param computer 
	 *            添加对象 
	 * @param cellStr 
	 *            单元格数据 
	 * @return 
	 */  
	private Computer addingComputer(int j, Computer computer, String cellStr) {  
		switch (j) {  
		case 0:  
			//computer.setId(0);  
			break;  
		case 1:  
			computer.setBrand(cellStr);  
			break;  
		case 2:  
			computer.setCpu(cellStr);  
			break;  
		case 3:  
			computer.setGpu(cellStr);  
			break;  
		case 4:  
			computer.setMemory(cellStr);  
			break;  
		case 5:  
			computer.setPrice(new Double(cellStr).doubleValue());  
			break;  
		}  

		return computer;  
	}  

	/** 
	 * 把单元格内的类型转换至String类型 
	 */  
	private String ConvertCellStr(Cell cell, String cellStr) {  

		switch (cell.getCellType()) {  

		case Cell.CELL_TYPE_STRING:  
			// 读取String  
			cellStr = cell.getStringCellValue().toString();  
			break;  

		case Cell.CELL_TYPE_BOOLEAN:  
			// 得到Boolean对象的方法  
			cellStr = String.valueOf(cell.getBooleanCellValue());  
			break;  

		case Cell.CELL_TYPE_NUMERIC:  

			// 先看是否是日期格式  
			if (DateUtil.isCellDateFormatted(cell)) {  

				// 读取日期格式  
				cellStr = cell.getDateCellValue().toString();  

			} else {  

				// 读取数字  
				cellStr = String.valueOf(cell.getNumericCellValue());  
			}  
			break;  

		case Cell.CELL_TYPE_FORMULA:  
			// 读取公式  
			cellStr = cell.getCellFormula().toString();  
			break;  
		}  
		return cellStr;  
	}
	/**
	 * 压缩导出
	 * @param list
	 * @param request
	 * @param length
	 * @param f
	 * @param out
	 * @throws IOException
	 */
	public void toExcel(List<Computer> list, HttpServletRequest request,  
			int length, String f, OutputStream out) throws IOException {  
		/**
		 *  response.setContentType("application/octet-stream;charset=UTF-8");  
            response.setHeader("Content-Disposition", "attachment;filename="  
                    + java.net.URLEncoder.encode(this.fileName, "UTF-8")  
                    + ".zip");  
            response.addHeader("Pargam", "no-cache");  
            response.addHeader("Cache-Control", "no-cache"); 

            Date date = new Date();  
        	SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");  
        	String f = "Person-" + format.format(date);  
        	this.fileName = f;  
        	setResponseHeader(response);  
        	OutputStream out = null;  
        	try {  
            	out = response.getOutputStream();  
            	List<Person> list = PersonService.getPerson();  
            	toExcel(list,request,10000,f,out);  
		 */
		List<String> fileNames = new ArrayList();// 用于存放生成的文件名称
		File zip = new File(request.getRealPath("/files") + "/" + f + ".zip");// 压缩文件  
		// 生成excel  
		for (int j = 0, n = list.size() / length + 1; j < n; j++) {  
			Workbook book = new HSSFWorkbook();  
			Sheet sheet = book.createSheet("Computer");  

			double d = 0;// 用来统计  
			String file = request.getRealPath("/files") + "/" + f + "-" + j  
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

				for (int i = 1, min = (list.size() - j * length + 1) > (length + 1) ? (length + 1)  
						: (list.size() - j * length + 1); i < min; i++) {  
					m++;  
					Computer user = list.get(length * (j) + i - 1);  
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
		FileInputStream inStream = new FileInputStream(zip);  
		byte[] buf = new byte[4096];  
		int readLength;  
		while (((readLength = inStream.read(buf)) != -1)) {  
			out.write(buf, 0, readLength);  
		}  
		inStream.close();  
	}  
}
