package com.execl.test.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelTest {




	public void test() throws FileNotFoundException, IOException {
		String pathName = "F:\\ceshi\\wirteExcel.xls";
		//1.创建一个工作簿Workbook
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		//2.创建一个工作表
		HSSFSheet sheet = workbook.createSheet();
		
		//3.创建行
		HSSFRow row = sheet.createRow(0);
		
		//4.创建单元格，设置值
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("张三");
		
		//5.写到一个文件中
		workbook.write(new FileOutputStream(pathName));
	}

}