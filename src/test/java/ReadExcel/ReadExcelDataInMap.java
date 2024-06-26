package ReadExcel;

import java.io.File;
import java.io.IOException;
import java.util.LinkedHashMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcelDataInMap {

    public static LinkedHashMap<String, String> getExcelDataAsMap(String excelFileName, String sheetName) throws EncryptedDocumentException, IOException {
        // Create a Workbook
        Workbook wb = WorkbookFactory.create(new File("src\\test\\resources\\excelFiles\\" + excelFileName + ".xlsx"));
        System.out.println("workbook===  "+ wb);
        // Get sheet with the given name "Sheet1"
        Sheet s = wb.getSheet(sheetName);
        System.out.println("excel sheet=== "+ s);

        // Initialized an empty LinkedHashMap which retain order
        LinkedHashMap<String, String> data = new LinkedHashMap<>();
        // Get total row count
        int rowCount = s.getPhysicalNumberOfRows();
        System.out.println("workbook rowCount === "+ rowCount);
        // Skipping first row as it contains headers
        for (int i = 1; i < rowCount; i++) {
            // Get the row
            Row r = s.getRow(i);
            System.out.println("skipping first row === "+ r);
            System.out.println("final worksheet === "+ s);
            // Since every row has two cells, first is field name and another is value.
            String fieldName = r.getCell(0).getStringCellValue();
            String fieldValue = r.getCell(1).getStringCellValue();
            data.put(fieldName, fieldValue);
        }
        return data;
    }

    public static void main(String[] args) throws EncryptedDocumentException, IOException {

        LinkedHashMap<String, String> mapData = getExcelDataAsMap("ExcelDataToReadInMap", "Sheet1");
        for (String s : mapData.keySet()) {
            System.out.println("-------------"+ s);
            System.out.println("-------------"+ mapData.keySet());
            System.out.println(mapData.get(s));
            System.out.println("Value of " + s + " is : " + mapData.get(s));
        }
    }

}
