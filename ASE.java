import java.io.*;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFRow;

public class ASE {

	public static void main(String[] args) {

		if(args.length < 2) {
			System.out.println("Usage : java -jar Access-Structure-Extractor.jar <access-database-path>.accdb/.mdb <path-to-excel-export>.xls");
			return;
		}

		String accessFile = args[0];
		String excelFile = args[1];

		HSSFWorkbook workbook = new HSSFWorkbook();

		try {
			Connection conn = DriverManager.getConnection("jdbc:ucanaccess://" + accessFile, "Admin", "");

			DatabaseMetaData md = conn.getMetaData();
			ResultSet rs = md.getTables(null, null, "%", null);

			while (rs.next()) {
				ResultSet rs2 = md.getColumns(null, null, rs.getString(3), "%");

				HSSFSheet sheet = workbook.createSheet(rs.getString(3));
				HSSFRow rowhead = sheet.createRow((short)0);
				rowhead.createCell(0).setCellValue("Nom");
				rowhead.createCell(1).setCellValue("Type");

				sheet.setColumnWidth(0, 10000);
				sheet.setColumnWidth(1, 6000);

				makeRowBold(workbook, rowhead);

				for (int i = 1; rs2.next(); i++) {
					HSSFRow row = sheet.createRow((short)i);
		            row.createCell(0).setCellValue(rs2.getString(4));
		            row.createCell(1).setCellValue(rs2.getString(6));
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}

		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(excelFile);
			workbook.write(fileOut);
	        fileOut.close();
	        System.out.println("Excel file generated !");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void makeRowBold(Workbook wb, HSSFRow row){
	    CellStyle style = wb.createCellStyle();
	    Font font = wb.createFont();
	    font.setBold(true);
	    style.setFont(font);

	    for(int i = 0; i < row.getLastCellNum(); i++) {
	    	row.getCell(i).setCellStyle(style);
	    }
	}

}
