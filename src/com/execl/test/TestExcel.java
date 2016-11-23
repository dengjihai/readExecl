package com.execl.test;

import com.execl.test.util.ExcelUtil;

import java.util.List;

public class TestExcel {

	/**
	 * execl 路径， 将接口文档中的响应字段复制出来到execl中
	 */
	static String mexeclPath="D:/bean.xlsx";


	public static void main(String[] args) throws  Exception {
		resp(mexeclPath);
		//req(mexeclPath);
	}

	/**
	 * 生成接口响应实体类
	 * @title: req
	 * @description: TODO(这里用一句话描述这个方法的作用)
	 * @param @param excelPath
	 * @param @throws Exception 设定文件
	 * @return void 返回类型
	 * @throws
	 */
	public static void req(String excelPath) throws Exception{
		List<List<List<String>>> fileInfo = ExcelUtil.readExcelWithoutTitle(excelPath);
		for (List<List<String>> list : fileInfo) {
			for (List<String> list2 : list) {
				String str = list2.toString();
				str = str.substring(1, str.length() - 1);
				String strs[] = str.split(",");
				if (!strs[0].isEmpty()) {
					String str1 = "\""+strs[0]+"\"";    
					System.out.println("@Query("+str1.trim()+") " + switchDataType(strs[1]) + " " + strs[0]+",");
				}
			}
		}
	}

	/**
	 *  
	* @title: resp  生成响应实体类
	* @description: TODO(这里用一句话描述这个方法的作用) 
	* @param @param excelPath execl文件路径
	* @param @throws Exception 设定文件 
	* @return void 返回类型 
	* @throws
	 */
	public static void resp(String excelPath) throws Exception {
		List<List<List<String>>> fileInfo = ExcelUtil
				.readExcelWithoutTitle(excelPath);
		for (List<List<String>> list : fileInfo) {
			for (List<String> list2 : list) {
				String str = list2.toString();
				str = str.substring(1, str.length() - 1);
				String strs[] = str.split(",");
				String dx = strs[1].substring(1, 2);
				String houmian = strs[1].substring(2, strs[1].length());
				System.out.println(" private " + dx + houmian + " " + strs[0]+ "; // " + strs[3]);
			}
		}
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
