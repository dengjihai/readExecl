package com.execl.test.util;

/**
 * 主运行类
 */
public class ReadExcelTest {


	static ReadExcelUtils excelUtils = new ReadExcelUtils();

	public static void main(String[] args) throws  Exception {
		//需要读取文件的路径
		String pathName = "D:\\bean.xls";
		/**
		 *  type 是响应实体类(resp)，还是请求字段类(req)
		 */
		excelUtils.readExcel(pathName,"resp");


	}


	

}

