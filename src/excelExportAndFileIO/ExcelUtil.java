package excelExportAndFileIO;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

	public static ResultSet rs = null;
	public static ResultSet SetFile(String WorkBook, String strSelectQuery, String strExcelPath) {
		rs = null;
		
		try {
			Connection con;
			con = DriverManager
					.getConnection("jdbc:odbc:"
							+ "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
							+ "DBQ=" + strExcelPath + ";" + "ReadOnly=0;");
			Statement s = con.createStatement(
					ResultSet.TYPE_SCROLL_INSENSITIVE,
					ResultSet.CONCUR_READ_ONLY);
			rs = s.executeQuery(strSelectQuery);
			
			if (!rs.isBeforeFirst() ) {    
				System.err.println("----NO SUIT SELECTED FOR EXECUTION----"); 
				 System.exit(0);
				} 
			
		} catch (SQLException ex) {
			ex.printStackTrace();
		}
		return rs;

	}
	
	public static String getcolumnvalue(ResultSet rss1, String colname) throws SQLException 
	{		
		   String colvalues = null;
		 	//System.out.println("Values:"+rs.getString(colname));		   	
        	colvalues = rss1.getString(colname);
            return colvalues;
	       
        
	}
	
	public static int getcount(ResultSet rss) throws SQLException 
	{
		   rss.next();
		   int count = rss.getInt("rowcount");
		   System.out.println("This Excel have " + count + " active row(s)");
		   return count;
	}

}


