package differentwaytotry.newoption;

import java.sql.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToAccess {

    // Method to import Excel data into MS Access
    public static void importExcelToAccess(String excelFilePath, String tableName) throws Exception {
        // Establish connection to MS Access DB using UCanAccess
        String dbURL = "jdbc:ucanaccess://C:/path/to/database.accdb";
        Connection conn = DriverManager.getConnection(dbURL);

        // Read Excel file
        FileInputStream fis = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        String sqlInsert = "INSERT INTO " + tableName + " (Column1, Column2, Column3) VALUES (?, ?, ?)";
        PreparedStatement pstmt = conn.prepareStatement(sqlInsert);

        // Iterate through Excel rows and insert into Access DB
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            pstmt.setString(1, row.getCell(0).getStringCellValue());
            pstmt.setString(2, row.getCell(1).getStringCellValue());
            pstmt.setString(3, row.getCell(2).getStringCellValue());
            pstmt.executeUpdate();
        }

        workbook.close();
        fis.close();
        pstmt.close();
        conn.close();
    }
}
