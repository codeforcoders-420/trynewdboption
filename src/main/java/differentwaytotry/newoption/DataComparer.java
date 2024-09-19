package differentwaytotry.newoption;

public class DataComparer {

    public static void compareTables() throws SQLException {
        String dbURL = "jdbc:ucanaccess://C:/path/to/database.accdb";
        Connection conn = DriverManager.getConnection(dbURL);

        Statement stmt = conn.createStatement();

        // Exact match query
        String exactMatchQuery = "INSERT INTO ExactMatch (Column1, Column2, Column3) " +
                                 "SELECT L.Column1, L.Column2, L.Column3 FROM LastWeekfile L " +
                                 "INNER JOIN ThisWeekfile T ON L.ID = T.ID " +
                                 "WHERE L.Column1 = T.Column1 AND L.Column2 = T.Column2";
        stmt.executeUpdate(exactMatchQuery);

        // Mismatch query
        String mismatchQuery = "INSERT INTO Mismatch (LastWeekColumn1, LastWeekColumn2, ThisWeekColumn1, ThisWeekColumn2) " +
                               "SELECT L.Column1, L.Column2, T.Column1, T.Column2 FROM LastWeekfile L " +
                               "INNER JOIN ThisWeekfile T ON L.ID = T.ID " +
                               "WHERE L.Column1 <> T.Column1 OR L.Column2 <> T.Column2";
        stmt.executeUpdate(mismatchQuery);

        // New rows query
        String newRowsQuery = "INSERT INTO NewRows (Column1, Column2, Column3) " +
                              "SELECT T.Column1, T.Column2, T.Column3 FROM ThisWeekfile T " +
                              "LEFT JOIN LastWeekfile L ON L.ID = T.ID " +
                              "WHERE L.ID IS NULL";
        stmt.executeUpdate(newRowsQuery);

        stmt.close();
        conn.close();
    }
    
    private void processFiles(File lastWeekFile, File thisWeekFile) {
        try {
            // Import Last Week file into LastWeekfile table
            ExcelToAccess.importExcelToAccess(lastWeekFile.getPath(), "LastWeekfile");
            
            // Import This Week file into ThisWeekfile table
            ExcelToAccess.importExcelToAccess(thisWeekFile.getPath(), "ThisWeekfile");

            // Compare tables and populate ExactMatch, Mismatch, and NewRows tables
            DataComparer.compareTables();

            System.out.println("Files processed successfully and data compared.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

