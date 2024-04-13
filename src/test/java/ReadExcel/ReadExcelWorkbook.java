package ReadExcel;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.LinkedHashMap;

public class ReadExcelWorkbook {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		/*
		 * It Creates the appropriate HSSFWorkbook / XSSFWorkbook from the given File,
		 * which must exist and be readable. In this case it will create a XSSFWorkbook
		 */
		Workbook wb = WorkbookFactory.create(new File("src\\test\\resources\\excelFiles\\MyTestData.xlsx"));
		// Get sheet with the given name "Sheet1"
		Sheet s = wb.getSheet("Sheet1");
		// Returns the number of physically defined rows (NOT the number of rows in the
		// sheet)
		int rowCount = s.getPhysicalNumberOfRows();
		System.out.println("total rows in sheet is : " + rowCount);

		// Initialized an empty LinkedHashMap which retain order
		LinkedHashMap<String, String> linkedHashMap = new LinkedHashMap<>();
		// Get total row count
		int wrowCount = s.getPhysicalNumberOfRows();
		System.out.println("workbook rowCount === "+ wrowCount);
		// Skipping first row as it contains headers
		for (int i = 1; i < rowCount; i++) {
			// Get the row
			Row r = s.getRow(i);
			System.out.println("skipping first row === "+ r);
			System.out.println("final worksheet === "+ s);
			// Since every row has two cells, first is field name and another is value.
			String fieldEmpID = r.getCell(0).getStringCellValue();
//			String EmpSalary = r.getCell(4).getStringCellValue();
//			String lastName = r.getCell(2).getStringCellValue();
//			linkedHashMap.put(fieldEmpID);

		}
		System.out.println("workbook data === "+ linkedHashMap);



	}
}
