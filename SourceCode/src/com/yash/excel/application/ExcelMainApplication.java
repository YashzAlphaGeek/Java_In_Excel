/**
 * 
 */
package com.yash.excel.application;

import com.yash.excel.manipulation.ReadingOfExcel;
import com.yash.excel.ui.ExcelChooserDialog;

/**
 * @author Yashwanth
 *
 */
public class ExcelMainApplication {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String filePath = ExcelChooserDialog.getFilePath();
		if (filePath != null) {
			ReadingOfExcel readExcel=new ReadingOfExcel();
			readExcel.fetchDataByColumnName(filePath);
			readExcel.printNetflixData();  
		}
	}

}
