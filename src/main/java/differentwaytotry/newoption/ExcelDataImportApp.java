package differentwaytotry.newoption;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import javafx.application.Application;
import javafx.stage.Stage;
import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.geometry.Insets;
import javafx.geometry.Pos;

import java.io.FileInputStream;
import java.io.FileWriter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;


import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.util.List;
import java.util.ArrayList;
import java.util.List;
import java.sql.ResultSet;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Objects;

import java.io.File;

public class ExcelDataImportApp extends Application {

	private static final String DATABASE_URL = "jdbc:ucanaccess://C:\\Users\\rajas\\Desktop\\Database\\filecompare.accdb";
	private static final String ACCESS_PATH = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\MSACCESS.EXE";
	private static final String DB_PATH = "C:\\Users\\rajas\\Desktop\\Database\\filecompare.accdb";
	private static final String SCRUB_RULES_FILE = "C:\\Users\\rajas\\Desktop\\Database\\Scrub rules.xlsx"; // Update with actual path
	private static final List<String> EXPECTED_HEADERS = Arrays.asList("Proc_code", "CMSAdd", "CMSTerm", "Modifiers", "Service", 
            "Service desc", "RateType", "Pricing Method", 
            "Rate Eff", "Rate Term", "MAxFee");

	private TextField lastWeekFilePathField;
	private TextField currentWeekFilePathField;

	public static void main(String[] args) {
		launch(args);
	}

	@Override
	public void start(Stage primaryStage) {
		primaryStage.setTitle("Select Excel Files for Last Week and Current Week");

		// Layout
		VBox layout = new VBox(10);
		layout.setPadding(new Insets(20, 20, 20, 20));

		// Last week file selection
		Label lastWeekLabel = new Label("Last Week Excel File:");
		lastWeekFilePathField = new TextField();
		lastWeekFilePathField.setEditable(false);
		lastWeekFilePathField.setPrefWidth(300);
		Button chooseLastWeekFileButton = new Button("Choose Last Week File");
		chooseLastWeekFileButton.setOnAction(e -> openFileChooser(primaryStage, lastWeekFilePathField));

		// Current week file selection
		Label currentWeekLabel = new Label("Current Week Excel File:");
		currentWeekFilePathField = new TextField();
		currentWeekFilePathField.setEditable(false);
		currentWeekFilePathField.setPrefWidth(300);
		Button chooseCurrentWeekFileButton = new Button("Choose Current Week File");
		chooseCurrentWeekFileButton.setOnAction(e -> openFileChooser(primaryStage, currentWeekFilePathField));

		// Submit button
		Button submitButton = new Button("Submit");
		submitButton.setOnAction(e -> submitFilePaths());

		// Adding components to layout
		layout.getChildren().addAll(lastWeekLabel, lastWeekFilePathField, chooseLastWeekFileButton, currentWeekLabel,
				currentWeekFilePathField, chooseCurrentWeekFileButton, submitButton);

		// Scene setup
		Scene scene = new Scene(layout, 400, 300);
		primaryStage.setScene(scene);
		primaryStage.show();
	}

	private void openFileChooser(Stage stage, TextField filePathField) {
		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("Select Excel File");
		fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx", "*.xls"));

		// Show open file dialog
		File selectedFile = fileChooser.showOpenDialog(stage);
		if (selectedFile != null) {
			filePathField.setText(selectedFile.getAbsolutePath());
		}
	}
	
	 private void updateHeadersIfNeeded(String filePath) {
	        try (FileInputStream fis = new FileInputStream(filePath);
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            Sheet sheet = workbook.getSheetAt(0);
	            Row headerRow = sheet.getRow(0);

	            if (headerRow == null) {
	                headerRow = sheet.createRow(0);
	            }

	            // Check each header and update if necessary
	            for (int i = 0; i < EXPECTED_HEADERS.size(); i++) {
	                Cell cell = headerRow.getCell(i);
	                if (cell == null || !EXPECTED_HEADERS.get(i).equalsIgnoreCase(cell.getStringCellValue().trim())) {
	                    // Update the cell with the expected header
	                    cell = headerRow.createCell(i);
	                    cell.setCellValue(EXPECTED_HEADERS.get(i));
	                }
	                sheet.autoSizeColumn(i);
	            }

	            // Save the updated Excel file to ensure headers are correct
	            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
	                workbook.write(fos);
	                System.out.println("Headers verified and updated if needed.");
	            }
	        } catch (Exception e) {
	            e.printStackTrace();
	            showAlert("Header Update Error", "An error occurred while verifying or updating the headers.");
	        }
	    }

	private void submitFilePaths() {
		String lastWeekFilePath = lastWeekFilePathField.getText();
		String currentWeekFilePath = currentWeekFilePathField.getText();

		if (lastWeekFilePath.isEmpty() || currentWeekFilePath.isEmpty()) {
			showAlert("Missing File", "Please select both the Last Week and Current Week files before submitting.");
		} else {
			try {
				updateHeadersIfNeeded(lastWeekFilePath);
                updateHeadersIfNeeded(currentWeekFilePath);
                
				runVBAScript(lastWeekFilePath, currentWeekFilePath);
				runComparisonQueriesnew();
				applyScrubRules(readScrubRules(SCRUB_RULES_FILE));
				exportUniqueRowsToExcel();
				
			} catch (IOException | InterruptedException e) {
				showAlert("Error", "An error occurred while running the VBA script.");
				e.printStackTrace();
			}
		}
	}
	
	 // Read scrub rules from Excel
    private List<ScrubRule> readScrubRules(String filePath) {
        List<ScrubRule> scrubRules = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from row 1 to skip headers
                Row row = sheet.getRow(i);
                if (row == null) continue;

                ScrubRule rule = new ScrubRule(
                        getCellValue(row.getCell(2)), // Service
                        getCellValue(row.getCell(3)), // RateType
                        getCellValue(row.getCell(5)), // Pricing Method
                        getCellValue(row.getCell(6)), // MaxFee
                        getCellValue(row.getCell(1))  // Rule Description
                );

                scrubRules.add(rule);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return scrubRules;
    }
    
 // Helper method to get cell value as a String
    private String getCellValue(Cell cell) {
        if (cell == null) return null;
        return cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : cell.toString();
    }

    // Apply scrub rules to tables
    private void applyScrubRules(List<ScrubRule> scrubRules) {
        try (Connection conn = DriverManager.getConnection(DATABASE_URL)) {
            for (ScrubRule rule : scrubRules) {
                // Apply rule to ExactMatch, MismatchLvC, and MismatchCvL tables
                //applyRuleToTable(conn, "ExactMatch", rule);
                applyRuleToTable(conn, "MismatchLvC", rule);
                applyRuleToTable(conn, "MismatchCvL", rule);
            }
            System.out.println("Scrub rules applied successfully.");
            showAlert("Scrub Rules Applied", "Scrub rules have been applied to the tables.");
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private void applyRuleToTable(Connection conn, String tableName, ScrubRule rule) throws SQLException {
        StringBuilder query = new StringBuilder("UPDATE " + tableName + " SET ScrubRule = 'Yes', [ScrubRuleDesc] = ? WHERE 1=1");
        
        List<String> params = new ArrayList<>();
        params.add(rule.RuleDescrption);

        if (rule.Service != null && !rule.Service.isEmpty()) {
            query.append(" AND Service = ?");
            params.add(rule.Service);
        }
        if (rule.RateType != null && !rule.RateType.isEmpty()) {
            query.append(" AND RateType = ?");
            params.add(rule.RateType);
        }
        if (rule.PricingMethod != null && !rule.PricingMethod.isEmpty()) {
            query.append(" AND [Pricing Method] = ?");
            params.add(rule.PricingMethod);
        }
        if (rule.MAxFee != null && !rule.MAxFee.isEmpty()) {
            query.append(" AND MAxFee = ?");
            params.add(rule.MAxFee);
        }
        
        System.out.println("Before Query to update is : " + query);

        try (PreparedStatement stmt = conn.prepareStatement(query.toString())) {
            for (int i = 0; i < params.size(); i++) {
                stmt.setString(i + 1, params.get(i));
                System.out.println("After Query to update is : " + stmt.toString());
                
            }
            stmt.executeUpdate();
        }
    }

    // ScrubRule class to store rule data
    static class ScrubRule {
        String Service;
        String RateType;
        String PricingMethod;
        String MAxFee;
        String RuleDescrption;

        ScrubRule(String service, String rateType, String pricingMethod, String maxfee, String ruleDescription) {
            this.Service = service;
            this.RateType = rateType;
            this.PricingMethod = pricingMethod;
            this.MAxFee = maxfee;
            this.RuleDescrption = ruleDescription;
        }
    }

	private void runComparisonQueries() {
		try (Connection conn = DriverManager.getConnection(DATABASE_URL); Statement stmt = conn.createStatement()) {

			// Create ExactMatch, Mismatch LvC, and Mismatch CvL tables if they don’t exist
			// createResultTables(stmt);

			// Insert exact matches
			String exactMatchQuery = "INSERT INTO ExactMatch\r\n" + "    SELECT lw.*\r\n"
					+ "    FROM Lastweekdata AS lw\r\n" + "    INNER JOIN CurrentWeekData AS cw\r\n"
					+ "    ON lw.Proc_code = cw.Proc_code \r\n" + "       AND lw.CMSAdd = cw.CMSAdd \r\n"
					+ "       AND lw.CMSTerm = cw.CMSTerm \r\n" + "       AND lw.Modifiers = cw.Modifiers \r\n"
					+ "       AND lw.Service = cw.Service \r\n"
					+ "       AND lw.[Service desc] = cw.[Service desc] \r\n"
					+ "       AND lw.RateType = cw.RateType \r\n"
					+ "       AND lw.[Pricing Method] = cw.[Pricing Method] \r\n"
					+ "       AND lw.[Rate Eff] = cw.[Rate Eff] \r\n"
					+ "       AND lw.[Rate Term] = cw.[Rate Term] \r\n" + "       AND lw.MAxFee = cw.MAxFee";
			stmt.executeUpdate(exactMatchQuery);
			System.out.println("Exact match records inserted into ExactMatch table.");
			showAlert("Validation Complete", "Exact matches and mismatches have been identified and saved.");

		} catch (SQLException e) {
			e.printStackTrace();
			showAlert("Error", "An error occurred while running the validation queries.");
		}
	}

	private void runComparisonQueriesnew() {
		try (Connection conn = DriverManager.getConnection(DATABASE_URL); Statement stmt = conn.createStatement()) {

			// stmt.executeUpdate("DELETE FROM ExactMatch");
			stmt.executeUpdate("DELETE FROM MismatchLvC");
			stmt.executeUpdate("DELETE FROM MismatchCvL");

			System.out.println("Tables cleared successfully.");

			// Fetch records from Lastweekdata and CurrentWeekData
			List<Record> lastweekRecords = fetchRecords(conn, "Lastweekdata");
			List<Record> currentweekRecords = fetchRecords(conn, "CurrentWeekData");

			PreparedStatement insertMismatchLvC = conn.prepareStatement(
					"INSERT INTO MismatchLvC (Proc_code, CMSAdd, CMSTerm, Modifiers, Service, [Service desc], RateType, [Pricing Method], [Rate Eff], [Rate Term], MAxFee, Differences) "
							+ "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
			PreparedStatement insertMismatchCvL = conn.prepareStatement(
					"INSERT INTO MismatchCvL (Proc_code, CMSAdd, CMSTerm, Modifiers, Service, [Service desc], RateType, [Pricing Method], [Rate Eff], [Rate Term], MAxFee, Differences) "
							+ "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");

			for (Record lw : lastweekRecords) {
				boolean exactMatch = false;
				StringBuilder differences = new StringBuilder();

				for (Record cw : currentweekRecords) {
					if (Objects.equals(lw.procCode, cw.procCode)) {
						if (!Objects.equals(lw.cmsAdd, cw.cmsAdd))
							differences.append("CMSAdd: ").append(lw.cmsAdd).append(" | ").append(cw.cmsAdd)
									.append("; ");
						if (!Objects.equals(lw.cmsTerm, cw.cmsTerm))
							differences.append("CMSTerm: ").append(lw.cmsTerm).append(" | ").append(cw.cmsTerm)
									.append("; ");
						if (!Objects.equals(lw.modifiers, cw.modifiers))
							differences.append("Modifiers: ").append(lw.modifiers).append(" | ").append(cw.modifiers)
									.append("; ");
						if (!Objects.equals(lw.service, cw.service))
							differences.append("Service: ").append(lw.service).append(" | ").append(cw.service)
									.append("; ");
						if (!Objects.equals(lw.serviceDesc, cw.serviceDesc))
							differences.append("Service desc: ").append(lw.serviceDesc).append(" | ")
									.append(cw.serviceDesc).append("; ");
						if (!Objects.equals(lw.rateType, cw.rateType))
							differences.append("RateType: ").append(lw.rateType).append(" | ").append(cw.rateType)
									.append("; ");
						if (!Objects.equals(lw.pricingMethod, cw.pricingMethod))
							differences.append("Pricing Method: ").append(lw.pricingMethod).append(" | ")
									.append(cw.pricingMethod).append("; ");
						if (!Objects.equals(lw.rateEff, cw.rateEff))
							differences.append("Rate Eff: ").append(lw.rateEff).append(" | ").append(cw.rateEff)
									.append("; ");
						if (!Objects.equals(lw.rateTerm, cw.rateTerm))
							differences.append("Rate Term: ").append(lw.rateTerm).append(" | ").append(cw.rateTerm)
									.append("; ");
						if (!Objects.equals(lw.maxFee, cw.maxFee))
							differences.append("MAxFee: ").append(lw.maxFee).append(" | ").append(cw.maxFee);

						// Check if there were no differences
						if (differences.length() == 0) {
							exactMatch = true;
						}
						break;
					}
				}

				if (differences.length() > 0) {
					// Insert into MismatchLvC with differences
					insertMismatchLvC.setString(1, lw.procCode);
					insertMismatchLvC.setString(2, lw.cmsAdd);
					insertMismatchLvC.setString(3, lw.cmsTerm);
					insertMismatchLvC.setString(4, lw.modifiers);
					insertMismatchLvC.setString(5, lw.service);
					insertMismatchLvC.setString(6, lw.serviceDesc);
					insertMismatchLvC.setString(7, lw.rateType);
					insertMismatchLvC.setString(8, lw.pricingMethod);
					insertMismatchLvC.setString(9, lw.rateEff);
					insertMismatchLvC.setString(10, lw.rateTerm);
					insertMismatchLvC.setString(11, lw.maxFee);
					insertMismatchLvC.setString(12, differences.toString());
					insertMismatchLvC.executeUpdate();
				}
			}

			for (Record cw : currentweekRecords) {
				boolean foundMatch = false;
				StringBuilder cwdifferences = new StringBuilder();

				for (Record lw : lastweekRecords) {
					if (Objects.equals(lw.procCode, cw.procCode)) {
						foundMatch = true; // Record found in Lastweekdata, check for differences

						if (!Objects.equals(cw.cmsAdd, lw.cmsAdd))
							cwdifferences.append("CMSAdd: ").append(cw.cmsAdd).append(" | ").append(lw.cmsAdd)
									.append("; ");
						if (!Objects.equals(cw.cmsTerm, lw.cmsTerm))
							cwdifferences.append("CMSTerm: ").append(cw.cmsTerm).append(" | ").append(lw.cmsTerm)
									.append("; ");
						if (!Objects.equals(cw.modifiers, lw.modifiers))
							cwdifferences.append("Modifiers: ").append(cw.modifiers).append(" | ").append(lw.modifiers)
									.append("; ");
						if (!Objects.equals(cw.service, lw.service))
							cwdifferences.append("Service: ").append(cw.service).append(" | ").append(lw.service)
									.append("; ");
						if (!Objects.equals(cw.serviceDesc, lw.serviceDesc))
							cwdifferences.append("Service desc: ").append(cw.serviceDesc).append(" | ")
									.append(lw.serviceDesc).append("; ");
						if (!Objects.equals(cw.rateType, lw.rateType))
							cwdifferences.append("RateType: ").append(cw.rateType).append(" | ").append(lw.rateType)
									.append("; ");
						if (!Objects.equals(cw.pricingMethod, lw.pricingMethod))
							cwdifferences.append("Pricing Method: ").append(cw.pricingMethod).append(" | ")
									.append(lw.pricingMethod).append("; ");
						if (!Objects.equals(cw.rateEff, lw.rateEff))
							cwdifferences.append("Rate Eff: ").append(cw.rateEff).append(" | ").append(lw.rateEff)
									.append("; ");
						if (!Objects.equals(cw.rateTerm, lw.rateTerm))
							cwdifferences.append("Rate Term: ").append(cw.rateTerm).append(" | ").append(lw.rateTerm)
									.append("; ");
						if (!Objects.equals(cw.maxFee, lw.maxFee))
							cwdifferences.append("MAxFee: ").append(cw.maxFee).append(" | ").append(lw.maxFee);

						break; // Break once the match is found and differences are recorded
					}

				}

				if (!foundMatch || cwdifferences.length() > 0) {
					// Insert into MismatchCvL with differences if no match or differences found
					insertMismatchCvL.setString(1, cw.procCode);
					insertMismatchCvL.setString(2, cw.cmsAdd);
					insertMismatchCvL.setString(3, cw.cmsTerm);
					insertMismatchCvL.setString(4, cw.modifiers);
					insertMismatchCvL.setString(5, cw.service);
					insertMismatchCvL.setString(6, cw.serviceDesc);
					insertMismatchCvL.setString(7, cw.rateType);
					insertMismatchCvL.setString(8, cw.pricingMethod);
					insertMismatchCvL.setString(9, cw.rateEff);
					insertMismatchCvL.setString(10, cw.rateTerm);
					insertMismatchCvL.setString(11, cw.maxFee);
					insertMismatchCvL.setString(12, cwdifferences.toString());
					insertMismatchCvL.executeUpdate();
				}
			}

			System.out.println("Mismatches identified and inserted into tables.");
			showAlert("Validation Complete", "Mismatches have been identified and saved.");

		} catch (SQLException e) {
			e.printStackTrace();
			showAlert("Error", "An error occurred while running the validation queries.");
		}
	}
	
	private void exportUniqueRowsToExcel() {
	    String[] columns = { "Proc_code", "Modifiers", "Rate Eff", "Rate Term", "MAxFee", "Differences" };

	    // Generate a timestamp for the file name
	    String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
	    String filePath = "C:\\Users\\rajas\\Desktop\\Database\\ChangeReport\\ChangeFile_OutputReport_" + timestamp + ".xlsx";

	    String query = "SELECT DISTINCT Proc_code, Modifiers, [Rate Eff], [Rate Term], MAxFee, Differences " +
	                   "FROM MismatchLvC WHERE ScrubRule IS NULL " +
	                   "UNION ALL " +
	                   "SELECT DISTINCT Proc_code, Modifiers, [Rate Eff], [Rate Term], MAxFee, Differences " +
	                   "FROM MismatchCvL WHERE ScrubRule IS NULL";

	    try (Connection conn = DriverManager.getConnection(DATABASE_URL);
	         Statement stmt = conn.createStatement();
	         ResultSet rs = stmt.executeQuery(query);
	         Workbook workbook = new XSSFWorkbook()) {

	        Sheet sheet = workbook.createSheet("Unique Changes");
	        Row headerRow = sheet.createRow(0);

	        // Creating header row
	        for (int i = 0; i < columns.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(columns[i]);
	            sheet.autoSizeColumn(i);
	        }

	        int rowNum = 1;
	        while (rs.next()) {
	            Row row = sheet.createRow(rowNum++);
	            row.createCell(0).setCellValue(rs.getString("Proc_code"));
	            row.createCell(1).setCellValue(rs.getString("Modifiers"));
	            row.createCell(2).setCellValue(rs.getString("Rate Eff"));
	            row.createCell(3).setCellValue(rs.getString("Rate Term"));
	            row.createCell(4).setCellValue(rs.getString("MAxFee"));
	            row.createCell(5).setCellValue(rs.getString("Differences"));
	        }

	        // Save the Excel file with a timestamp in the file name
	        try (FileOutputStream fileOut = new FileOutputStream(new File(filePath))) {
	            workbook.write(fileOut);
	            System.out.println("Excel report generated and saved at: " + filePath);
	            showAlert("Export Complete", "The ChangeFile_OutputReport.xlsx file has been saved to the shared path.");
	        }
	    } catch (SQLException e) {
	        e.printStackTrace();
	        showAlert("Database Error", "An error occurred while querying the database.");
	    } catch (Exception e) {
	        e.printStackTrace();
	        showAlert("File Error", "An error occurred while creating the Excel report.");
	    }
	}


	// Helper method to fetch records
	private List<Record> fetchRecords(Connection conn, String tableName) throws SQLException {
		List<Record> records = new ArrayList<>();
		String query = "SELECT * FROM " + tableName;
		try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery(query)) {
			while (rs.next()) {
				records.add(new Record(rs.getString("Proc_code"), rs.getString("CMSAdd"), rs.getString("CMSTerm"),
						rs.getString("Modifiers"), rs.getString("Service"), rs.getString("Service desc"),
						rs.getString("RateType"), rs.getString("Pricing Method"), rs.getString("Rate Eff"),
						rs.getString("Rate Term"), rs.getString("MAxFee")));
			}
		}
		return records;
	}

	// Record class definition
	static class Record {
		String procCode;
		String cmsAdd;
		String cmsTerm;
		String modifiers;
		String service;
		String serviceDesc;
		String rateType;
		String pricingMethod;
		String rateEff;
		String rateTerm;
		String maxFee;

		public Record(String procCode, String cmsAdd, String cmsTerm, String modifiers, String service,
				String serviceDesc, String rateType, String pricingMethod, String rateEff, String rateTerm,
				String maxFee) {
			this.procCode = procCode;
			this.cmsAdd = cmsAdd;
			this.cmsTerm = cmsTerm;
			this.modifiers = modifiers;
			this.service = service;
			this.serviceDesc = serviceDesc;
			this.rateType = rateType;
			this.pricingMethod = pricingMethod;
			this.rateEff = rateEff;
			this.rateTerm = rateTerm;
			this.maxFee = maxFee;
		}
	}

	private void runVBAScript(String lastWeekFilePath, String currentWeekFilePath)
			throws IOException, InterruptedException {
		// Path to save the generated VBA script
		String vbaScriptPath = "C:\\Users\\rajas\\Desktop\\Database\\ImportScript.vbs"; // Set the path to save the VBA
																						// script

		// Generate the VBA script with the provided file paths
		generateVBAScript(vbaScriptPath, lastWeekFilePath, currentWeekFilePath);

		// Execute the VBA script using the Windows Script Host
		Process process = Runtime.getRuntime().exec("wscript " + vbaScriptPath);
		process.waitFor();

		if (process.exitValue() == 0) {
			showAlert("Success", "Data imported successfully from both Excel files.");
		} else {
			showAlert("Error", "An error occurred during the import process.");
		}
	}

	private void generateVBAScript(String vbaScriptPath, String lastWeekFilePath, String currentWeekFilePath)
			throws IOException {
		// VBA script code
		String vbaScript = "Set objAccess = CreateObject(\"Access.Application\")\n"
				+ "objAccess.OpenCurrentDatabase \"C:\\Users\\rajas\\Desktop\\Database\\filecompare.accdb\"\n"
				+ "objAccess.Run \"ImportExcelFiles\", \"" + lastWeekFilePath + "\", \"" + currentWeekFilePath + "\"\n"
				+ "objAccess.Quit\n" + "Set objAccess = Nothing";

		// Write the VBA script to the specified file
		try (FileWriter writer = new FileWriter(vbaScriptPath)) {
			writer.write(vbaScript);
			System.out.println("VBA script created successfully at " + vbaScriptPath);
		}
	}
	
	private void exportExactMatchQueryToExcel() {
        try {
            String vbaCommand = MSACCESS_PATH + " \"" + DATABASE_PATH + "\" /x ExportExactMatchQueryToExcel";
            Process p = Runtime.getRuntime().exec(vbaCommand);
            p.waitFor();
            System.out.println("ExactMatchQuery results have been exported to Excel.");
            showAlert("Export Complete", "The exact match query results have been exported to Excel.");
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
            showAlert("Error", "An error occurred while exporting the query results to Excel.");
        }
    }

    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

	private void showAlert(String title, String message) {
		Alert alert = new Alert(Alert.AlertType.INFORMATION);
		alert.setTitle(title);
		alert.setHeaderText(null);
		alert.setContentText(message);
		alert.showAndWait();
	}

}