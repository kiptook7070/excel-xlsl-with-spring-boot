package ReadExcel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcelDataInMap2 {

	public static LinkedHashMap<String, List<String>> getExcelDataAsMap(String excelFileName, String sheetName) throws EncryptedDocumentException, IOException {
		// Create a Workbook
		Workbook wb = WorkbookFactory.create(new File("src\\test\\resources\\excelFiles\\"+excelFileName+".xlsx"));
		// Get sheet with the given name "Sheet1"
		Sheet s = wb.getSheet(sheetName);
		// Initialized an empty LinkedHashMap which retain order
		LinkedHashMap<String, List<String>> data = new LinkedHashMap<>();
		// Get total row count
		int rowCount = s.getPhysicalNumberOfRows();
		// Skipping first row as it contains headers
		for (int i = 1; i < rowCount; i++) {
			// Get the row
			Row r = s.getRow(i);
			// Since every row has two cells, first is field name and another is value.
			String fieldName = r.getCell(0).getStringCellValue();
			String fieldValue = r.getCell(1).getStringCellValue();
			if(data.containsKey(fieldName))
			{
				List<String> existingValues = data.get(fieldName);
				existingValues.add(fieldValue);
				data.put(fieldName, existingValues);
			}
			else
			{
				List<String> newValues = new ArrayList<>();;
				newValues.add(fieldValue);
				data.put(fieldName, newValues);
			}
		}
		return data;
	}

	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		LinkedHashMap<String,List<String>> mapData = getExcelDataAsMap("ExcelDataToReadInMapDuplicateKeys","Sheet1");
		for(String s: mapData.keySet())
		{
			System.out.println("Value of "+s +" is : "+mapData.get(s));
		}
	}

}
