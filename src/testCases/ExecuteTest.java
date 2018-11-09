package testCases;

import java.io.File;
import java.io.InputStream;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import operation.ReadObject;
import operation.UIOperation;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.Test;

import java.io.IOException;

import excelExportAndFileIO.ExcelUtil;
import excelExportAndFileIO.GenericFunctions;

@Test
public class ExecuteTest {

	public String CurrSuiteDescription;
	public String Suite_Excel;
	public static String strTestCaseID;
	public List<String> strExcelSuitPath = new ArrayList<String>();
	public File f ;
	WebDriver webdriver = new FirefoxDriver();
	
	public void SuiteDetails() throws Exception {
		String strSelectQuery;
		int Suite_counter = 0;
	
		String strExcelPath = "D:\\Style_Automation\\Test Cases\\Driver.xls";
		f = new File(strExcelPath);
		if (! f.isFile())
		{
		System.err.println("Driver file does not exists at "  + strExcelPath);	
		System.exit(0);
		}
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
		strSelectQuery = GenericFunctions.GetSelectQuery("Execution", "Count");
		ResultSet res = ExcelUtil.SetFile("Execution", strSelectQuery,
				strExcelPath);
		Suite_counter = ExcelUtil.getcount(res);
		res.close();
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////		
		strSelectQuery = GenericFunctions.GetSelectQuery("Execution",
				"ExecutionWorkBook");
		res = ExcelUtil.SetFile("Execution", strSelectQuery, strExcelPath);
		res.beforeFirst();
		
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
		for (int iSuiteRowCounter = 0; iSuiteRowCounter < Suite_counter; iSuiteRowCounter++) {
			res.next();

			CurrSuiteDescription = ExcelUtil.getcolumnvalue(res, "Description");
			System.out.println("++++++++++++++++++++ SUIT LEVEL DESCRIPTION +++++++++++++++++++");
			Suite_Excel = ExcelUtil.getcolumnvalue(res, "Suite_Workbook");

			System.out.println("Suite_Excels That Needs To Be Executed :: - " + Suite_Excel);
			System.out.println("Suite_Description :: - "
					+ CurrSuiteDescription);
			strExcelSuitPath.add("D:\\Style_Automation\\Test Cases\\"
					+ Suite_Excel);
			System.out.println("Suite_Excels Path That Will Be Used To Check Modules:: - " + "D:\\Style_Automation\\Test Cases\\"
					+ Suite_Excel );
		}
		res.close();

	} /// Close Function

	public void TestDetails() throws Exception {
		int Suite_Excel_counter = strExcelSuitPath.size();
		UIOperation operation = new UIOperation(webdriver);
		
		for (int Suite_Excel_Iterator = 0; Suite_Excel_Iterator < Suite_Excel_counter; Suite_Excel_Iterator++) {
			f = new File(strExcelSuitPath.get(Suite_Excel_Iterator));
			if (f.isFile())
			{
			String strSelectQuery = GenericFunctions.GetSelectQuery(
					"Execution", "Count");
			ResultSet res = ExcelUtil.SetFile("Execution", strSelectQuery,
					strExcelSuitPath.get(Suite_Excel_Iterator));
			int TestCases_ModuleCounter = ExcelUtil.getcount(res);
			res.close();
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
			strSelectQuery = GenericFunctions.GetSelectQuery("Execution",
					"ExecutionWorkBook");
			res = ExcelUtil.SetFile("Execution", strSelectQuery,
					strExcelSuitPath.get(Suite_Excel_Iterator));  /// Find the Modules That needs to be run from Excel given in Suit_WorkBook
			res.beforeFirst();  
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
			for (int TestCases_Module_Iterator = 0; TestCases_Module_Iterator < TestCases_ModuleCounter; TestCases_Module_Iterator++) {
				
				System.out.println("++++++++++++++++++++ Module LEVEL DESCRIPTION +++++++++++++++++++");
				System.out.println("Checking " +strExcelSuitPath.get(Suite_Excel_Iterator)+ " Excel Path for Vaild Modules ");
				res.next();
				
				String Module = ExcelUtil.getcolumnvalue(res, "Module");
				System.out.println("System will Execute Module :: - " + Module);
				String TestCase_Name = ExcelUtil.getcolumnvalue(res, "TestCase_Name");
				System.out.println(Module + " TestCase_Name is :: - " + TestCase_Name);
				strTestCaseID = ExcelUtil.getcolumnvalue(res, "TestCaseNo");
				System.out.println(Module + " TestCaseNo :: - " + strTestCaseID);
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////			
				strSelectQuery = GenericFunctions.GetSelectQuery(Module,"TCCount");
// 				Module ="Flight";
				ResultSet Wb_res = ExcelUtil.SetFile(Module, strSelectQuery,
						strExcelSuitPath.get(Suite_Excel_Iterator));
				int Module_TC_Counter = ExcelUtil.getcount(Wb_res);
				Wb_res.close();
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////						
				System.out.println("+++++++++++++Test Case Starting +++++++++++++++++++");
				strSelectQuery = GenericFunctions.GetSelectQuery(Module,"TestCaseSheet");
				Wb_res = ExcelUtil.SetFile(Module, strSelectQuery,
						strExcelSuitPath.get(Suite_Excel_Iterator));
				Wb_res.beforeFirst();
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
				for (int Module_TC_iterator = 0; Module_TC_iterator < Module_TC_Counter; Module_TC_iterator++) {
					Wb_res.next();

					String strTS_Id, strTestStepDetail, strKeyword, strIdentifier, strObjectType, strTestData_1, strTestData_2, DictKeyword, strStringToBeEvaluated;
					String blnActionStatus, strKey;

					strTS_Id = ExcelUtil.getcolumnvalue(Wb_res, "TSID");
					System.out.println("+++++++++++++Test Case " + strTS_Id + " Details +++++++++++++++++++");
					System.out.println("Test Case Id is :: " + strTS_Id);
					strTestStepDetail = ExcelUtil.getcolumnvalue(Wb_res, "TestStep");
					System.out.println("TestStep is :: " + strTS_Id);
					strKeyword = ExcelUtil.getcolumnvalue(Wb_res, "Keyword");
					System.out.println("Keyword is :: " + strKeyword);
					strObjectType = ExcelUtil.getcolumnvalue(Wb_res, "ByObjectType");
					System.out.println("Object Type is :: " + strObjectType);
					strIdentifier = ExcelUtil.getcolumnvalue(Wb_res, "Identifier");
					System.out.println("Identifier is :: " + strIdentifier);
					strTestData_1 = ExcelUtil.getcolumnvalue(Wb_res, "TestData_1");
					System.out.println("TestData_1 is :: " + strTestData_1);
					strTestData_2 = ExcelUtil.getcolumnvalue(Wb_res, "TestData_2");
					System.out.println("TestData_2 is :: " + strTestData_2); 

					//Call perform function to perform operation on UI
	    	    	operation.perform(strKeyword, strObjectType, strIdentifier, strTestData_1);
					
				} //Module_TC_iterator For loop
			  } // TestCases_Module_Iterator For loop
			}
			else {
		        System.err.println("Suit files does not exits " + strExcelSuitPath.get(Suite_Excel_Iterator));
		        System.exit(0);
		    }	// If Condition
		}// Suite_Excel_Iterator For loop
	  } // TestDetails function closing
} // Class closing

/*
 * 
 * For iSuiteRowCounter = 0 to AllSuitesExcelObject.GetRowCount - 1 Dim
 * strCurrSuiteSheet,SuiteExecutionExcel,strSuiteWorkbookPath,iTcRowCounter
 * 
 * Environment.Value("CurrSuiteName") =
 * AllSuitesExcelObject.GetCurrentCellValue("Suite_Name") strCurrSuiteSheet =
 * AllSuitesExcelObject.GetCurrentCellValue("Suite_Workbook")
 * strSuiteWorkbookPath = strTestCaseFolderPath & "\" & strCurrSuiteSheet Print
 * Environment.Value("CurrSuiteName") Print strCurrSuiteSheet
 * 
 * If IfFileExists(strSuiteWorkbookPath) Then Set SuiteExecutionExcel =
 * Fn_NewExcelReader() SuiteExecutionExcel.SetFile
 * strSuiteWorkbookPath,GetSelectQuery(strExecutionSheetName,"ExecutionSheet")
 * 
 * For iTcRowCounter = 0 to SuiteExecutionExcel.GetRowCount - 1 Dim
 * strModuleName,TestCaseExcel,iTsRowCounter,strTestCaseStatus
 * 
 * Environment.Value("TestCaseID") =
 * SuiteExecutionExcel.GetCurrentCellValue("TestCaseNo")
 * Environment.Value("TestCaseName") =
 * SuiteExecutionExcel.GetCurrentCellValue("TestCase_Name")
 * Environment.value("TestCaseFullName") = Environment.Value("CurrSuiteName") &
 * "." & Environment.Value("TestCaseName") cp_SetResultLog
 * Environment.Value("TestCaseFullName"
 * ),SuiteExecutionExcel.GetCurrentCellValue("Description") strTestCaseStatus =
 * "Pass"
 * 
 * strModuleName = SuiteExecutionExcel.GetCurrentCellValue("Module") Set
 * TestCaseExcel = Fn_NewExcelReader() TestCaseExcel.SetFile
 * strSuiteWorkbookPath,GetSelectQuery(strModuleName,"TestCaseSheet")
 * 
 * RemoveAllvaluesFromDict
 * 
 * For iTsRowCounter = 0 to TestCaseExcel.GetRowCount-1 Dim
 * strTS_Id,strTestStepDetail
 * ,strKeyword,strField,strTestData_1,strTestData_2,DictKeyword
 * ,strStringToBeEvaluated Dim blnActionStatus, strKey
 * 
 * strTS_Id = TestCaseExcel.GetCurrentCellValue("TSID") Print strTS_Id
 * strTestStepDetail = TestCaseExcel.GetCurrentCellValue("TestStep") strKeyword
 * = TestCaseExcel.GetCurrentCellValue("Keyword") strField =
 * TestCaseExcel.GetCurrentCellValue("Field") strTestData_1 =
 * TestCaseExcel.GetCurrentCellValue("TestData_1") strTestData_2 =
 * TestCaseExcel.GetCurrentCellValue("TestData_2")
 * 
 * If Lcase(Left(strTestData_1,Len(strDataPrefix))) = Lcase(strDataPrefix) Then
 * strKey=Mid(strTestData_1,Len(strDataPrefix)+1)
 * strTestData_1=GetDataValueFromDict(Lcase(strkey)) End If If not
 * (Trim(strTestData_2) ="") And Lcase(Left(strTestData_2,Len(strDataPrefix))) =
 * Lcase(strDataPrefix) Then strKey = Mid(strTestData_2,Len(strDataPrefix)+1)
 * strTestData_2=GetDataValueFromDict(Lcase(strkey)) End If
 * 
 * Set DictKeyword = CreateObject("Scripting.Dictionary") DictKeyword.Add
 * "Field",strField DictKeyword.Add "TestData_1",strTestData_1 DictKeyword.Add
 * "TestData_2",strTestData_2 DictKeyword.Add "TestStepID",strTS_Id
 * DictKeyword.Add "TestStepDetails",strTestStepDetail
 * 
 * strStringToBeEvaluated = "blnActionStatus  =" & " " & strKeyWord &
 * "( DictKeyword )" Execute (strStringToBeEvaluated)
 * 
 * If not Cbool(blnActionStatus) Then strTestCaseStatus = "Fail" Msgbox "Fail"
 * Exit For End If TestCaseExcel.MoveToNextRecord Next cp_EndReport
 * Environment.Value("TestCaseFullName"),strTestCaseStatus
 * SuiteExecutionExcel.MoveToNextRecord Next Else msgbox
 * "Suit files does not exits ["& strSuiteWorkbookPath &"]" End If
 * 
 * AllSuitesExcelObject.MoveToNextRecord Next
 */

