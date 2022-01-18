/**
 * 
 */
package com.yash.excel.manipulation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Yashwanth
 *
 */
public class ReadingOfExcel {

	private Map<String,String> netflixMap= new HashMap<String, String>();

	public void readExcel()
	{

	}

	/**
	 * Fetching of Data from Excel Sheet
	 * 
	 * @param filePath - Excel File Path
	 * @return excel column data - Block With Port and their Tag Values
	 */
	public void fetchDataByColumnName(String filePath) {
		Map<String, Integer> headerColumnMap = new HashMap<>();
		Workbook workbook = null;
		try (FileInputStream inputStream = new FileInputStream(new File(filePath))) {
			workbook = new XSSFWorkbook(inputStream);
			Sheet sheet = workbook.getSheet("Sheet1");
			sheet.getRow(0).forEach(cell -> {
				headerColumnMap.put(cell.getStringCellValue(), cell.getColumnIndex());
			});
			getColumnData("Netflix Series", "Synopsis", sheet, headerColumnMap, workbook);
		}
		catch (IOException e) {
			e.printStackTrace();
		}

	}

	/**
	 * @param netflixSeriesCol
	 * @param synopsisColumn
	 * @param sheet
	 * @param headerColumnMap
	 * @param workbook
	 * @param netflixMap 
	 * @return
	 */
	private void getColumnData(String netflixSeriesCol, String synopsisColumn, Sheet sheet, Map<String, Integer> headerColumnMap,
			Workbook workbook) {
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			try {
				Cell netflixSeriesCell = sheet.getRow(i).getCell(headerColumnMap.get(netflixSeriesCol));
				Cell synopsisCell = sheet.getRow(i).getCell(headerColumnMap.get(synopsisColumn));
				if(netflixSeriesCell!=null && synopsisCell!=null)
				{
					String netflixSeriesName = getCellStringValue(workbook, netflixSeriesCell);
					String synopsisDetail = getCellStringValue(workbook, synopsisCell);	
					netflixMap.put(netflixSeriesName, synopsisDetail);
				}

			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
	}


	/**
	 * Get Cell String Value from Excel
	 * 
	 * @param workbook - Workbook
	 * @param cell - Excel Cell
	 */
	private String getCellStringValue(Workbook workbook, Cell cell) {
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		CellType cellTypeEnum = formulaEvaluator.evaluateInCell(cell).getCellTypeEnum();
		if (cellTypeEnum == CellType.STRING) {
			return cell.getStringCellValue();
		}
		return null;
	}

	/**
	 * @return 
	 * 
	 */
	public void printNetflixData() {
		for(Entry<String, String> netflixData:netflixMap.entrySet())
		{
			System.out.println(netflixData.getKey()+"==>"+netflixData.getValue());
		}
	}

}
