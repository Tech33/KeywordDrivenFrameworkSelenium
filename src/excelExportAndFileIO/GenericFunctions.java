package excelExportAndFileIO;

import testCases.ExecuteTest;

public class GenericFunctions extends ExecuteTest {

	public static String strExecutionFlagColmnName = "Execution_Flag";

	public static String GetSelectQuery(String strSheetName, String strSheetType) {
		String strSelectQuery = null;
		switch (strSheetType) {
		case "ExecutionWorkBook":
			strSelectQuery = ("Select * from [" + strSheetName + "$] Where "
					+ strExecutionFlagColmnName + " Like 'Yes%' OR "
					+ strExecutionFlagColmnName + " Like 'yes%'");
			System.out.println(strSelectQuery);
			break;
		case "TestCaseSheet":
			String strTestCaseIdColmn = "TCID";
			strSelectQuery = ("Select * from [" + strSheetName + "$] Where "
					+ strTestCaseIdColmn + " = '" + strTestCaseID + "'" );
			System.out.println(strSelectQuery);
			break;
		case "Count":
			strSelectQuery = ("SELECT COUNT(*) AS rowcount FROM ["
					+ strSheetName + "$] Where " + strExecutionFlagColmnName
					+ " Like 'Yes%' OR " + strExecutionFlagColmnName + " Like 'yes%'");
			System.out.println(strSelectQuery);
			break;
		case "TCCount":
			strTestCaseIdColmn = "TCID";
			strSelectQuery = ("SELECT COUNT(*) AS rowcount FROM ["
					+ strSheetName + "$] Where " + strTestCaseIdColmn
					+ " = '" + strTestCaseID + "'" );
			System.out.println(strSelectQuery);
			break;
		}
		return strSelectQuery;
	}
}
