package ReadExcel;

import java.io.File;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class UnitTests {

	@Test
	public void differentNumberOfSheets() throws EncryptedDocumentException, IOException
	{
		String userDir = System.getProperty("user.dir");
		Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithThreeSheets.xlsx"));
		Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithThreeSheets.xlsx"));
//		Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithTwoSheets-Data Actual.xlsx"));
		MSE_CompareExcelFiles mse_CompareExcelFiles = new MSE_CompareExcelFiles();
		mse_CompareExcelFiles.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
		
	}
	
	@Test
	public void sameNumberOfSheets() throws EncryptedDocumentException, IOException
	{
		String userDir = System.getProperty("user.dir");
		Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithThreeSheets.xlsx"));
		Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithThreeSheets.xlsx"));
		MSE_CompareExcelFiles mse_CompareExcelFiles = new MSE_CompareExcelFiles();
		mse_CompareExcelFiles.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
	}
	
	
	@Test
	public void differentNumberOfRows() throws EncryptedDocumentException, IOException
	{
		String userDir = System.getProperty("user.dir");
		Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithTwoSheets - Data Actual.xlsx"));
		Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithTwoSheets- Extra Row.xlsx"));
		MSE_CompareExcelFiles mse_CompareExcelFiles = new MSE_CompareExcelFiles();
		mse_CompareExcelFiles.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
	}
	
	@Test
	public void differentNumberOfColumns() throws EncryptedDocumentException, IOException
	{
		String userDir = System.getProperty("user.dir");
		Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithTwoSheets - Data Actual.xlsx"));
		Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithTwoSheets- Extra Columns.xlsx"));
		MSE_CompareExcelFiles mse_CompareExcelFiles = new MSE_CompareExcelFiles();
		mse_CompareExcelFiles.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
	}
	
	@Test
	public void sameContentOfExcelFiles() throws EncryptedDocumentException, IOException {
		String userDir = System.getProperty("user.dir");
		Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithThreeSheets.xlsx"));
		Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\excelFiles\\CompareExcel\\ExcelFilesWithThreeSheets.xlsx"));
		MSE_CompareExcelFiles mse_CompareExcelFiles = new MSE_CompareExcelFiles();
		mse_CompareExcelFiles.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
		mse_CompareExcelFiles.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
		mse_CompareExcelFiles.verifyDataInExcelBookAllSheets(wb1, wb2);
	
	}
	
	
}
