package sample;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.*;

/**
 * Created by Nikita Tkachuk
 */
public class ExcelWorker
{
	public void startParse(String sourceFilePath, ParserConfiguration configuration, String pathToSave)
	{
		writeFile(ExcelProductTemplate.INSTANCE.getTemplateSheet(), sourceFilePath, configuration, pathToSave);
	}

	private void writeFile(Sheet sheet, String sourceFileName, ParserConfiguration configuration, String pathToSave)
	{
		Map<Integer, List<Product>> productsByPortions = productsParser(sourceFileName, configuration);
		ExcelCopier copier = new ExcelCopier();
		for (Map.Entry<Integer, List<Product>> entry : productsByPortions.entrySet())
		{
			HSSFWorkbook workbook = new HSSFWorkbook();

			HSSFSheet newSheet = workbook.createSheet("Лист1");
			for (Product product : entry.getValue())
			{
				HSSFSheet templateSheet = (HSSFSheet) sheet;

				productToExcel(templateSheet, product);
				try
				{
					URL mainImageUrl = new URL(product.getImageUrl());
					String protocol = mainImageUrl.getProtocol();
					String host = mainImageUrl.getHost();
					List<URL> urlCollection = new ArrayList<>();
					for (String string : product.getModelRow())
					{
						urlCollection.add(new URL(protocol + "://" + host + string));
					}
					insertModelRowImages(templateSheet, urlCollection);
					insertMainImage(templateSheet, mainImageUrl);
				}
				catch (MalformedURLException e)
				{
					e.printStackTrace();
				}
				copier.addOldSheetToEnd(newSheet, templateSheet, true);
			}
			try (OutputStream outputStream = new FileOutputStream(new File(pathToSave + "\\_" + entry.getKey() + ".xls")))
			{
				workbook.write(outputStream);
			}
			catch (IOException e)
			{
				e.printStackTrace();
			}

		}
	}

	public static void productToExcel(HSSFSheet sheet, Product product)
	{
		Map<String, String> mappedProductOnCells = product.productToExcelCellMapping();
		for (Map.Entry<String, String> entry : mappedProductOnCells.entrySet())
		{
			CellReference ref = new CellReference(entry.getKey());
			Row r = sheet.getRow(ref.getRow());
			if (r != null)
			{
				Cell c = r.getCell(ref.getCol());
				if(c != null)
				{
					c.setCellValue(entry.getValue());
				}
			}
		}
	}

	private static void insertMainImage(HSSFSheet sheet, URL imageUrl)
	{
		try(InputStream urlStream = imageUrl.openStream())
		{
			byte [] bytes = IOUtils.toByteArray(urlStream);
			HSSFWorkbook wb = sheet.getWorkbook();
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
			CreationHelper helper = wb.getCreationHelper();
			Drawing drawing = sheet.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();
			anchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
			anchor.setCol1(0);
			anchor.setRow1(0);
			anchor.setCol2(4);
			anchor.setRow2(21);
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			pict.resize();
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	private static void insertModelRowImages(HSSFSheet sheet, List<URL> urlCollection)
	{
		int row1 = 0;
		int row2 = 0;
		int col = 12;  //L
		for (int i = 0; i < urlCollection.size(); i++)
		{
			URL imageUrl = urlCollection.get(i);
			try(InputStream urlStream = imageUrl.openStream())
			{
				byte [] bytes = IOUtils.toByteArray(urlStream);
				HSSFWorkbook wb = sheet.getWorkbook();
				int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
				CreationHelper helper = wb.getCreationHelper();
				Drawing drawing = sheet.createDrawingPatriarch();
				ClientAnchor anchor = helper.createClientAnchor();
				anchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
				if(i > 0 && i % 3 == 0)
				{
					col = col + 2;
					row1 = 0;
					row2 = 0;
				}
				row2 = row2 + 7;
				anchor.setRow1(row1);
				anchor.setCol1(col);
				anchor.setRow2(row2);
				anchor.setCol2(col + 2);
				Picture pict = drawing.createPicture(anchor, pictureIdx);
				row1 = row2;
				pict.resize();
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
		}

	}

	private Map<Integer, List<Product>> productsParser(String fileName, ParserConfiguration configuration)
	{
		Map<Integer, List<Product>> result = new HashMap<>();
		List<Product> onePortion = new ArrayList<>(100);
		try (InputStream stream = new FileInputStream(fileName))
		{
			HSSFWorkbook workbook = new HSSFWorkbook(stream);
			for (int i = 0; i < workbook.getNumberOfSheets(); i++)
			{
				if (configuration.ignoreList.contains(i))
				{
					continue;
				}
				HSSFSheet sheet = workbook.getSheetAt(i);

				boolean isTitleRow = false;
				for (Row currentRow : sheet)
				{
					if (!isTitleRow)
					{
						isTitleRow = true;
						continue;
					}
					Product product = new Product();
					Integer nameColumn = configuration.nameColumn;
					if (nameColumn != null)
					{
						Cell cell = currentRow.getCell(nameColumn);
						if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING)
						{
							product.setName(parseCellValue(cell));
						} else
						{
							continue;
						}
					}
					Integer articleColumn = configuration.articleColumn;
					if (articleColumn != null)
					{
						Cell cell = currentRow.getCell(articleColumn);
						if (cell != null)
						{
							product.setArticle(parseCellValue(cell));
						}
					}
					Integer colorColumn = configuration.colorColumn;
					if (colorColumn != null)
					{
						Cell cell = currentRow.getCell(colorColumn);
						if (cell != null)
						{
							product.setColor(parseCellValue(cell));
						}
					}
					Integer costColumn = configuration.costColumn;
					if (costColumn != null)
					{
						Cell cell = currentRow.getCell(costColumn);
						if (cell != null)
						{
							product.setCost(parseCellValue(cell));
						}
					}
					Integer inPortionColumn = configuration.inPortionColumn;
					if (inPortionColumn != null)
					{
						Cell cell = currentRow.getCell(inPortionColumn);
						if (cell != null)
						{
							String stringCellValue = parseCellValue(cell);
							if (stringCellValue.contains("шт"))
							{
								stringCellValue = stringCellValue.replaceAll("шт", "");
							}
							product.setPortionSize(stringCellValue);
						}
					}
					Integer descriptionColumn = configuration.descriptionColumn;
					if (descriptionColumn != null)
					{
						Cell cell = currentRow.getCell(descriptionColumn);
						if (cell != null)
						{
							product.setDescription(parseCellValue(cell));
						}
					}
					Integer imageColumn = configuration.imageColumn;
					if (imageColumn != null)
					{
						Cell cell = currentRow.getCell(imageColumn);
						if (cell != null)
						{
							product.setImageUrl(parseCellValue(cell));
						}
					}
					Integer modelRowColumn = configuration.modelRowColumn;
					if (modelRowColumn != null)
					{
						Cell cell = currentRow.getCell(modelRowColumn);
						if (cell != null)
						{
							product.setModelRow(parseModelRow(parseCellValue(cell)));
						}
					}
					List<Integer> sizeColumns = configuration.sizeColumns;
					if (sizeColumns != null && !sizeColumns.isEmpty())
					{
						StringBuilder builder = new StringBuilder();
						for (Integer sizeColumn : sizeColumns)
						{
							Cell cell = currentRow.getCell(sizeColumn);
							if (cell != null)
							{
								builder.append(parseCellValue(cell));
								builder.append(". ");
							}
						}
						if(configuration.delimiterSizes == null || "".equals(configuration.delimiterSizes))
						{
							product.setSizes(Collections.singletonList(builder.toString()));
						}
						else
						{
							product.setSizes(parseSizes(builder.toString(), configuration.delimiterSizes));
						}
					}
					onePortion.add(product);
					if (onePortion.size() == 100)
					{
						result.put(result.keySet().size(), onePortion);
						onePortion = new ArrayList<>(100);
					}
				}
			}
			if (!onePortion.isEmpty())
			{
				result.put(result.keySet().size(), onePortion);
			}

		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return result;
	}

	private List<String> parseSizes(String stringCellValue, String delimiter)
	{
		if (stringCellValue.contains(delimiter))
		{
			return Arrays.asList(stringCellValue.split(delimiter));
		}
		/*if (stringCellValue.contains(", "))
		{
			result.addAll(Arrays.asList(stringCellValue.split(", ")));
			return result;
		}
		if (stringCellValue.contains("Носки") && stringCellValue.contains("лет"))
		{
			result.add(stringCellValue);
			return result;
		}
		if (stringCellValue.contains("."))
		{
			result.addAll(Arrays.asList(stringCellValue.split(".")));
			return result;
		}
		if (stringCellValue.contains(","))
		{
			result.addAll(Arrays.asList(stringCellValue.split(",")));
			return result;
		}
		if (stringCellValue.contains("/"))
		{
			result.addAll(Arrays.asList(stringCellValue.split("/")));
			return result;
		}
		if (stringCellValue.contains("-"))
		{
			result.addAll(Arrays.asList(stringCellValue.split("-")));
			return result;
		}
		if (stringCellValue.contains(" "))
		{
			result.addAll(Arrays.asList(stringCellValue.split(" ")));
			return result;
		}*/

		return Collections.emptyList();
	}

	private List<String> parseModelRow(String modelRow)
	{
		if(!"".equals(modelRow))
		{
			return Arrays.asList(modelRow.split(";"));
		}
		else
		{
			return Collections.emptyList();
		}
	}

	private String parseCellValue(Cell cell)
	{
		switch (cell.getCellType())
		{
			case HSSFCell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			case HSSFCell.CELL_TYPE_NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			case HSSFCell.CELL_TYPE_BLANK:
				return "";
			case HSSFCell.CELL_TYPE_BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case HSSFCell.CELL_TYPE_ERROR:
				return String.valueOf(cell.getErrorCellValue());
			case HSSFCell.CELL_TYPE_FORMULA:
				return cell.getCellFormula();
			default:
				return "";
		}
	}
}
