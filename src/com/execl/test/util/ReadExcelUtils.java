package com.execl.test.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcelUtils {
	/**
	 * @param pathName
	 * @param  type 是响应实体类(resp)，还是请求字段类(req)
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public void readExcel(String pathName,String type) throws IOException,
			FileNotFoundException {
		//1.找到一个已经存在的excel(工作簿)
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(pathName));
		
		readSheetsFromWorkbook(workbook,type);
	}

	/**
	 * 读取所有的工作表
	 * @param workbook
	 */
	private void readSheetsFromWorkbook(HSSFWorkbook workbook,String type) {
		//2.选定一个工作表
		for(int sheetIndex = 0;sheetIndex<workbook.getNumberOfSheets();sheetIndex++){
			HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			System.out.println("工作表名称："+sheet.getSheetName());
			readRowFromSheet(sheet,type);
		}
	}

	/**
	 * 读取所有行
	 * @param sheet
	 */
	private void readRowFromSheet(HSSFSheet sheet,String type) {
		//3.选定行
		for(int rowIndex = 0;;rowIndex++){
			HSSFRow row = sheet.getRow(rowIndex);
			if(row==null){
				break;
			}
			String lineStr=readCellValueFromRow(row);
			//System.out.println(lineStr);
			if(type.equals("req")){
				generateReq(lineStr);
			}else{
				generateRsp(lineStr);
			}
		}
	}

	/**
	 * 读取一行的所有列
	 * @param row
	 */
	private String readCellValueFromRow(HSSFRow row) {
		StringBuffer sb= new StringBuffer();
		//4.选定列，单元格
		for(int cellnum = 0; ; cellnum++){
			HSSFCell cell = row.getCell(cellnum);
			if(cell==null){
				break;
			}
//			int cellType = cell.getCellType();
//			if(Cell.CELL_TYPE_NUMERIC==cellType){
//				double cellValue = cell.getNumericCellValue();
//				if(cellnum!=0){
//					System.out.print("\t");
//				}
//				System.out.print(cellValue);
//			}else if(Cell.CELL_TYPE_STRING==cellType){
//				String cellValue = cell.getStringCellValue();
//				if(cellnum!=0){
//					System.out.print("\t");
//				}
//				System.out.print(cellValue);
//			}
			
			//魔法方法：把各种类型统一转成string类型

			
			//5.根据单元格的类型调用不同的取值方法
			cell.setCellType(Cell.CELL_TYPE_STRING);
			String cellValue = cell.getStringCellValue();
			if(cellnum!=0){
				//System.out.print("\t");
			}
			sb.append(cellValue);
			sb.append(",");
			//System.out.print("=="+cellValue);
		}
		return  sb.toString();
	}

	public void generateReq(String cellValue){
		String strs[] = cellValue.split(",");
		if (!strs[0].isEmpty()) {
			String str1 = "\""+strs[0]+"\"";
			System.out.println("@Query("+str1.trim()+") " + switchDataType(strs[1]) + " " + strs[0]+",");
		}
	}

	public void generateRsp(String str){
		str = str.substring(0, str.length() - 1);
		String strs[] = str.split(",");
		System.out.println(" private " + switchDataType(strs[1])
				+ " " + strs[0]
				+ "; // " + strs[3]);
	}

	public static String switchDataType(String type) {
		String str = "String";
		if(type.contains("int")){
			str = "Integer";
		}
		else if(type.contains("string")){
			str = "String";
		}
		else if(type.contains("double")){
			str = "Double";
		}
		else if(type.contains("datetime")){
			str = "Date";
		}
		else if(type.contains("long")){
			str = "Long";
		}
		return str;
	}
}
