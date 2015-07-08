package sample;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        FXMLLoader loader  = new FXMLLoader(getClass().getClassLoader().getResource("sample.fxml"));
	    Parent root = (Parent)loader.load();
	    Controller controller = loader.getController();
	    controller.setStage(primaryStage);
        primaryStage.setTitle("Excel converter");
        primaryStage.setScene(new Scene(root, 550, 500));
	    primaryStage.getIcons().add(new Image(getClass().getClassLoader().getResourceAsStream("icon.png")));
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
