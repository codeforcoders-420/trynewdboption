package differentwaytotry.newoption;

import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;

import java.io.File;

public class ExcelFileUploader extends Application {

    private File lastWeekFile;
    private File thisWeekFile;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel File Uploader");

        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));

        Button uploadLastWeekButton = new Button("Upload Last Week File");
        Button uploadThisWeekButton = new Button("Upload This Week File");

        uploadLastWeekButton.setOnAction(e -> {
            lastWeekFile = fileChooser.showOpenDialog(primaryStage);
            if (lastWeekFile != null) {
                System.out.println("Selected Last Week File: " + lastWeekFile.getPath());
            }
        });

        uploadThisWeekButton.setOnAction(e -> {
            thisWeekFile = fileChooser.showOpenDialog(primaryStage);
            if (thisWeekFile != null) {
                System.out.println("Selected This Week File: " + thisWeekFile.getPath());
            }
        });

        Button processButton = new Button("Process Files");
        processButton.setOnAction(e -> {
            if (lastWeekFile != null && thisWeekFile != null) {
                // Call method to process and import files into MS Access
                processFiles(lastWeekFile, thisWeekFile);
            } else {
                System.out.println("Please upload both files.");
            }
        });

        VBox vbox = new VBox(10, uploadLastWeekButton, uploadThisWeekButton, processButton);
        Scene scene = new Scene(vbox, 400, 200);

        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void processFiles(File lastWeekFile, File thisWeekFile) {
        // Implement the logic to import files into MS Access DB and compare
        System.out.println("Processing files...");
    }

    public static void main(String[] args) {
        launch(args);
    }
}
