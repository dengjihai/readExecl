package com.execl.test.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {
	
	/**
	 * 适用于没有标题行的excel，例如
	 * 张三   25岁     男   175cm
	 * 李四   22岁     女   160cm
	 * 每一行构成一个map，key值是列标题，value是列值。没有值的单元格其value值为null
	 * 返回结果最外层的list对应一个excel文件，第二层的list对应一个sheet页，第三层的map对应sheet页中的一行
	 * @throws Exception 
	 */
	public static List<List<List<String>>> readExcelWithoutTitle(String filepath) throws Exception{
	    String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());
	    InputStream is = null;
	    Workbook wb = null;
	    try {
	        is = new FileInputStream(filepath);
	         
	        if (fileType.equals("xls")) {
	            wb = new HSSFWorkbook(is);
	        } else if (fileType.equals("xlsx")) {
	            wb = new XSSFWorkbook(is);
	        } else {
	            throw new Exception("读取的不是excel文件");
	        }
	         
	        List<List<List<String>>> result = new ArrayList<List<List<String>>>();//对应excel文件
	         
	        int sheetSize = wb.getNumberOfSheets();
	        for (int i = 0; i < sheetSize; i++) {//遍历sheet页
	            Sheet sheet = wb.getSheetAt(i);
	            List<List<String>> sheetList = new ArrayList<List<String>>();//对应sheet页
	             
	            int rowSize = sheet.getLastRowNum() + 1;
	            for (int j = 0; j < rowSize; j++) {//遍历行
	                Row row = sheet.getRow(j);
	                if (row == null) {//略过空行
	                    continue;
	                }
	                int cellSize = row.getLastCellNum();//行中有多少个单元格，也就是有多少列
	                List<String> rowList = new ArrayList<String>();//对应一个数据行
	                for (int k = 0; k < cellSize; k++) {
	                    Cell cell = row.getCell(k);
	                    String value = null;
	                    if (cell != null) {
	                        value = cell.toString();
	                    }
	                    rowList.add(value);
	                }
	                sheetList.add(rowList);
	            }
	            result.add(sheetList);
	        }
	         
	        return result;
	    } catch (FileNotFoundException e) {
	        throw e;
	    } finally {
	        if (wb != null) {
	            wb.close();
	        }
	        if (is != null) {
	            is.close();
	        }
	    }
	}
	
	
	/**
	 * 适用于第一行是标题行的excel，例如
	 * 姓名   年龄  性别  身高
	 * 张三   25  男   175
	 * 李四   22  女   160
	 * 每一行构成一个map，key值是列标题，value是列值。没有值的单元格其value值为null
	 * 返回结果最外层的list对应一个excel文件，第二层的list对应一个sheet页，第三层的map对应sheet页中的一行
	 * @throws Exception 
	 */
	public static List<List<Map<String, String>>> readExcelWithTitle(String filepath) throws Exception{
	    String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());
	    InputStream is = null;
	    Workbook wb = null;
	    try {
	        is = new FileInputStream(filepath);
	         
	        if (fileType.equals("xls")) {
	            wb = new HSSFWorkbook(is);
	        } else if (fileType.equals("xlsx")) {
	            wb = new XSSFWorkbook(is);
	        } else {
	            throw new Exception("读取的不是excel文件");
	        }
	         
	        List<List<Map<String, String>>> result = new ArrayList<List<Map<String, String>>>();//对应excel文件
	         
	        int sheetSize = wb.getNumberOfSheets();
	        for (int i = 0; i < sheetSize; i++) {//遍历sheet页
	            Sheet sheet = wb.getSheetAt(i);
	            List<Map<String, String>> sheetList = new ArrayList<Map<String, String>>();//对应sheet页
	             
	            List<String> titles = new ArrayList<String>();//放置所有的标题
	             
	            int rowSize = sheet.getLastRowNum() + 1;
	            for (int j = 0; j < rowSize; j++) {//遍历行
	                Row row = sheet.getRow(j);
	                if (row == null) {//略过空行
	                    continue;
	                }
	                int cellSize = row.getLastCellNum();//行中有多少个单元格，也就是有多少列
	                if (j == 0) {//第一行是标题行
	                    for (int k = 0; k < cellSize; k++) {
	                        Cell cell = row.getCell(k);
	                        titles.add(cell.toString());
	                    }
	                } else {//其他行是数据行
	                    Map<String, String> rowMap = new HashMap<String, String>();//对应一个数据行
	                    for (int k = 0; k < titles.size(); k++) {
	                        Cell cell = row.getCell(k);
	                        String key = titles.get(k);
	                        String value = null;
	                        if (cell != null) {
	                            value = cell.toString();
	                        }
	                        rowMap.put(key, value);
	                    }
	                    sheetList.add(rowMap);
	                }
	            }
	            result.add(sheetList);
	        }
	         
	        return result;
	    } catch (FileNotFoundException e) {
	        throw e;
	    } finally {
	        if (wb != null) {
	            wb.close();
	        }
	        if (is != null) {
	            is.close();
	        }
	    }
	}
	
	//默认单元格内容为数字时格式
	private static DecimalFormat df = new DecimalFormat("0");
	// 默认单元格格式化日期字符串 
	private static SimpleDateFormat sdf = new SimpleDateFormat(  "yyyy-MM-dd HH:mm:ss"); 
	// 格式化数字
	private static DecimalFormat nf = new DecimalFormat("0.00");  
	public static ArrayList<ArrayList<Object>> readExcel(File file){
		if(file == null){
			return null;
		}
		if(file.getName().endsWith("xlsx")){
			//处理ecxel2007
			return readExcel2007(file);
		}else{
			//处理ecxel2003
			return readExcel2003(file);
		}
	}
	/*
	 * @return 将返回结果存储在ArrayList内，存储结构与二位数组类似
	 * lists.get(0).get(0)表示过去Excel中0行0列单元格
	 */
	public static ArrayList<ArrayList<Object>> readExcel2003(File file){
		try{
			ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();
			ArrayList<Object> colList;
			HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
			HSSFSheet sheet = wb.getSheetAt(0);
			HSSFRow row;
			HSSFCell cell;
			Object value;
			for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
				row = sheet.getRow(i);
				colList = new ArrayList<Object>();
				if(row == null){
					//当读取行为空时
					if(i != sheet.getPhysicalNumberOfRows()){//判断是否是最后一行
						rowList.add(colList);
					}
					continue;
				}else{
					rowCount++;
				}
				for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
					cell = row.getCell(j);
					if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
						//当该单元格为空
						if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格
							colList.add("");
						}
						continue;
					}
					switch(cell.getCellType()){
					 case XSSFCell.CELL_TYPE_STRING:  
		                    System.out.println(i + "行" + j + " 列 is String type");  
		                    value = cell.getStringCellValue();  
		                    break;  
		                case XSSFCell.CELL_TYPE_NUMERIC:  
		                    if ("@".equals(cell.getCellStyle().getDataFormatString())) {  
		                        value = df.format(cell.getNumericCellValue());  
		                    } else if ("General".equals(cell.getCellStyle()  
		                            .getDataFormatString())) {  
		                        value = nf.format(cell.getNumericCellValue());  
		                    } else {  
		                        value = sdf.format(HSSFDateUtil.getJavaDate(cell  
		                                .getNumericCellValue()));  
		                    }  
		                    System.out.println(i + "行" + j  
		                            + " 列 is Number type ; DateFormt:"  
		                            + value.toString()); 
		                    break;  
		                case XSSFCell.CELL_TYPE_BOOLEAN:  
		                    System.out.println(i + "行" + j + " 列 is Boolean type");  
		                    value = Boolean.valueOf(cell.getBooleanCellValue());
		                    break;  
		                case XSSFCell.CELL_TYPE_BLANK:  
		                    System.out.println(i + "行" + j + " 列 is Blank type");  
		                    value = "";  
		                    break;  
		                default:  
		                    System.out.println(i + "行" + j + " 列 is default type");  
		                    value = cell.toString();  
					}// end switch
					colList.add(value);
				}//end for j
				rowList.add(colList);
			}//end for i
			
			return rowList;
		}catch(Exception e){
			return null;
		}
	}
	
	public static ArrayList<ArrayList<Object>> readExcel2007(File file){
		try{
			ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();
			ArrayList<Object> colList;
			XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
			XSSFSheet sheet = wb.getSheetAt(0);
			XSSFRow row;
			XSSFCell cell;
			Object value;
			for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
				row = sheet.getRow(i);
				colList = new ArrayList<Object>();
				if(row == null){
					//当读取行为空时
					if(i != sheet.getPhysicalNumberOfRows()){//判断是否是最后一行
						rowList.add(colList);
					}
					continue;
				}else{
					rowCount++;
				}
				for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
					cell = row.getCell(j);
					if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
						//当该单元格为空
						if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格
							colList.add("");
						}
						continue;
					}
					switch(cell.getCellType()){
					 case XSSFCell.CELL_TYPE_STRING:  
		                    System.out.println(i + "行" + j + " 列 is String type");  
		                    value = cell.getStringCellValue();  
		                    break;  
		                case XSSFCell.CELL_TYPE_NUMERIC:  
		                    if ("@".equals(cell.getCellStyle().getDataFormatString())) {  
		                        value = df.format(cell.getNumericCellValue());  
		                    } else if ("General".equals(cell.getCellStyle()  
		                            .getDataFormatString())) {  
		                        value = nf.format(cell.getNumericCellValue());  
		                    } else {  
		                        value = sdf.format(HSSFDateUtil.getJavaDate(cell  
		                                .getNumericCellValue()));  
		                    }  
		                    System.out.println(i + "行" + j  
		                            + " 列 is Number type ; DateFormt:"  
		                            + value.toString()); 
		                    break;  
		                case XSSFCell.CELL_TYPE_BOOLEAN:  
		                    System.out.println(i + "行" + j + " 列 is Boolean type");  
		                    value = Boolean.valueOf(cell.getBooleanCellValue());
		                    break;  
		                case XSSFCell.CELL_TYPE_BLANK:  
		                    System.out.println(i + "行" + j + " 列 is Blank type");  
		                    value = "";  
		                    break;  
		                default:  
		                    System.out.println(i + "行" + j + " 列 is default type");  
		                    value = cell.toString();  
					}// end switch
					colList.add(value);
				}//end for j
				rowList.add(colList);
			}//end for i
			
			return rowList;
		}catch(Exception e){
			System.out.println("exception");
			return null;
		}
	}
	
	public static void writeExcel(ArrayList<ArrayList<Object>> result,String path){
		if(result == null){
			return;
		}
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("sheet1");
		for(int i = 0 ;i < result.size() ; i++){
			 HSSFRow row = sheet.createRow(i);
			if(result.get(i) != null){
				for(int j = 0; j < result.get(i).size() ; j ++){
					HSSFCell cell = row.createCell(j);
					cell.setCellValue(result.get(i).get(j).toString());
				}
			}
		}
		ByteArrayOutputStream os = new ByteArrayOutputStream();
        try
        {
            wb.write(os);
        } catch (IOException e){
            e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        File file = new File(path);//Excel文件生成后存储的位置。
        OutputStream fos  = null;
        try
        {
            fos = new FileOutputStream(file);
            fos.write(content);
            os.close();
            fos.close();
        }catch (Exception e){
            e.printStackTrace();
        }           
	}
	
	public static DecimalFormat getDf() {
		return df;
	}
	public static void setDf(DecimalFormat df) {
		ExcelUtil.df = df;
	}
	public static SimpleDateFormat getSdf() {
		return sdf;
	}
	public static void setSdf(SimpleDateFormat sdf) {
		ExcelUtil.sdf = sdf;
	}
	public static DecimalFormat getNf() {
		return nf;
	}
	public static void setNf(DecimalFormat nf) {
		ExcelUtil.nf = nf;
	}
	
	
}