package by.dabl;

import by.dabl.model.ExcelWorker;
import by.dabl.model.ParserConfiguration;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class Controller
{

	public Button openFileButton;
	public TextField absoluteFilePathName;
	public TextField absoluteFolderPath;

	public TextField nameColumn;
	public TextField articleColumn;
	public TextField colorColumn;
	public TextField costColumn;
	public TextField inPortionColumn;
	public TextField sizeColumn;
	public TextField descriptionColumn;
	public TextField ignoreLists;
	public TextField imageColumn;
	public TextField modelRow;
	public TextField sizesDelimiter;
	private Stage stage;


	public Controller()
	{

	}

	public Stage getStage()
	{
		return stage;
	}

	public void setStage(Stage stage)
	{
		this.stage = stage;
	}

	@FXML
	public void openFileDialog()
	{
		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("Выберите файл для парсинга");
		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
		fileChooser.getExtensionFilters().add(extFilter);
		File file = fileChooser.showOpenDialog(stage);
		if(file != null)
		{
			absoluteFilePathName.setText(file.getAbsolutePath());
		}
	}

	@FXML
	public void openFolderToSave()
	{
		DirectoryChooser fileChooser = new DirectoryChooser();
		fileChooser.setTitle("Выберите папку для сохраненя");
		File file = fileChooser.showDialog(stage);
		absoluteFolderPath.setText(file.getAbsolutePath());
	}

	public void starParsing()
	{
		Task<Void> task = new Task<Void>()
		{
			@Override
			protected Void call() throws Exception
			{
				ExcelWorker worker = new ExcelWorker(absoluteFilePathName.getText(), createConfiguration(), absoluteFolderPath.getText());
				Thread workerThread = new Thread(worker, "Worker Thread");
				workerThread.start();
				//worker.run();
				return null;
			}
		};
		task.progressProperty();
		new Thread(task, "Task thread").start();
	}

	private ParserConfiguration createConfiguration()
	{
		ParserConfiguration configuration = new ParserConfiguration();
		String emptyString = "";
		String nameColumnText = nameColumn.getText();
		String articleColumnText = articleColumn.getText();
		String costColumnText = costColumn.getText();
		String colorColumnText = colorColumn.getText();
		String inPortionColumnText = inPortionColumn.getText();
		String descriptionColumnText = descriptionColumn.getText();
		String imageColumnText = imageColumn.getText();
		String modelRowColumnText = modelRow.getText();
		String sizesDelimiterText = sizesDelimiter.getText();
		if(!emptyString.equals(nameColumnText))
		{
			configuration.nameColumn = Integer.valueOf(nameColumnText);
		}
		if(!emptyString.equals(articleColumnText))
		{
			configuration.articleColumn = Integer.valueOf(articleColumnText);
		}
		if(!emptyString.equals(costColumnText))
		{
			configuration.costColumn = Integer.valueOf(costColumnText);
		}
		if(!emptyString.equals(colorColumnText))
		{
			configuration.colorColumn = Integer.valueOf(colorColumnText);
		}
		if(!emptyString.equals(inPortionColumnText))
		{
			configuration.inPortionColumn = Integer.valueOf(inPortionColumnText);
		}
		if(!emptyString.equals(descriptionColumnText))
		{
			configuration.descriptionColumn = Integer.valueOf(descriptionColumnText);
		}
		if(!emptyString.equals(imageColumnText))
		{
			configuration.imageColumn = Integer.valueOf(imageColumnText);
		}
		if(!emptyString.equals(modelRowColumnText))
		{
			configuration.modelRowColumn = Integer.valueOf(modelRowColumnText);
		}
		if(!emptyString.equals(sizesDelimiterText))
		{
			configuration.delimiterSizes = sizesDelimiterText;
		}
		configuration.sizeColumns = parseSizeColumns();
		configuration.ignoreList = parseIgnoreList();
		return configuration;
	}

	private List<Integer> parseIgnoreList()
	{
		String ignoreListText = ignoreLists.getText();
		if("".equals(ignoreListText))
		{
			return Collections.emptyList();
		}
		if(ignoreListText.contains(","))
		{
			List<Integer> result = new ArrayList<>();
			String [] ignoredListsAsStrings = ignoreListText.split(",");
			for (String listForIgnore : ignoredListsAsStrings)
			{
				result.add(Integer.valueOf(listForIgnore));
			}
			return result;
		}
		else
		{
			return Collections.singletonList(Integer.valueOf(ignoreListText));
		}

	}

	private List<Integer> parseSizeColumns()
	{
		String sizes = sizeColumn.getText();
		if("".equals(sizes))
		{
			return Collections.emptyList();
		}
		if(sizes.contains(","))
		{
			List<Integer> result = new ArrayList<>();
			Integer start = Integer.valueOf(sizes.split(",")[0]);
			Integer end = Integer.valueOf(sizes.split(",")[1]);
			while (start < end)
			{
				result.add(start);
				start++;
			}
			return result;
		}
		else
		{
			return Collections.singletonList(Integer.valueOf(sizes));
		}
	}


}
