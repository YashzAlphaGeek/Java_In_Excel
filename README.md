# Java_In_Excel

When working with large sets of data that need any level of automation, we strongly recommend using Java to create, edit, compute and download associated Excel files. For most of the functions, you only need few lines of code.

Let us assume we have large sets of Netflix Data in Excel Sheet and we would like to fetch the data column wise by giving the column name.

<kbd> <img src="https://github.com/YashzAlphaGeek/Java_In_Excel/blob/master/Images/ExcelData.png"/> </kbd>

## Reading of Excel by Column Name

<pre><code>

public void fetchDataByColumnName(String filePath) {
		Map<String, Integer> headerColumnMap = new HashMap<>();
		Workbook workbook = null;
		try (FileInputStream inputStream = new FileInputStream(new File(filePath))) {
			workbook = new XSSFWorkbook(inputStream);
			Sheet sheet = workbook.getSheet("Sheet1");
			<b>sheet.getRow(0).forEach(cell -> {
				headerColumnMap.put(cell.getStringCellValue(), cell.getColumnIndex());
			});
			getColumnData("Netflix Series", "Synopsis", sheet, headerColumnMap, workbook);</b>
		}
		catch (IOException e) {
			e.printStackTrace();
		}
	}


private void getColumnData(String netflixSeriesCol, String synopsisColumn, Sheet sheet, Map<String, Integer> headerColumnMap,Workbook workbook) {
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			try {
				Cell netflixSeriesCell = sheet.getRow(i).getCell(headerColumnMap.get(netflixSeriesCol));
				Cell synopsisCell = sheet.getRow(i).getCell(headerColumnMap.get(synopsisColumn));
				if(netflixSeriesCell!=null && synopsisCell!=null)
				{
					<b>String netflixSeriesName = getCellStringValue(workbook, netflixSeriesCell);</b>
					<b>String synopsisDetail = getCellStringValue(workbook, synopsisCell);</b>
					netflixMap.put(netflixSeriesName, synopsisDetail);
				}

			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
	}  
</code></pre>

 ## Final Outcome:
 
 <kbd> <img src="https://github.com/YashzAlphaGeek/Java_In_Excel/blob/master/Images/ReadColumnByName.png"/> </kbd>

------------------------------------------------------------------------------------
“Thanks for watching. If you liked this page, make sure to subscribe for more!”

	First, solve the problem. Then, write the code. 
------------------------------------------------------------------------------------
:grinning:
