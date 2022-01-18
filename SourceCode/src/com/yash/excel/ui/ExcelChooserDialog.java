/**
 * 
 */
package com.yash.excel.ui;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 * @author Yashwanth
 *
 */
public class ExcelChooserDialog {


	/**
	 * Excel File Path
	 */
	public static String excelFilePath;


	/**
	 * @return the excelFilePath
	 */
	public static String getExcelFilePath() {
		return excelFilePath;
	}


	/**
	 * @param excelFilePath the excelFilePath to set
	 */
	public static void setExcelFilePath(final String excelFilePath) {
		ExcelChooserDialog.excelFilePath = excelFilePath;
	}


	/**
	 * File Chooser Dialog
	 *
	 * @return excel file path
	 */
	private static String fileChooserDialog() {
		JFileChooser excelChooser = new JFileChooser();
		JFrame chooserFrame = new JFrame();
		excelChooser.setDialogTitle("Please choose the excel sheet");
		FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
		excelChooser.setFileFilter(filter);
		// ChooserFrame is set to the front
		chooserFrame.toFront();
		chooserFrame.setAlwaysOnTop(true);
		chooserFrame.setAutoRequestFocus(true);
		int returnValue = excelChooser.showOpenDialog(chooserFrame);
		try {
			if (returnValue == JFileChooser.APPROVE_OPTION) {
				ExcelChooserDialog.excelFilePath = excelChooser.getSelectedFile().getPath();
				if (null != ExcelChooserDialog.excelFilePath) {
					setExcelFilePath(ExcelChooserDialog.excelFilePath);
					chooserFrame.dispose();
				}
			}
		}

		catch (Exception e) {
			e.printStackTrace();
		}
		return ExcelChooserDialog.excelFilePath;
	}


	/**
	 * Get the filepath
	 *
	 * @return file chooser path
	 */
	public static String getFilePath() {
		return fileChooserDialog();
	}

}
