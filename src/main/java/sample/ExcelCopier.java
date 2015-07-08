package sample;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.*;

public final class ExcelCopier
{
	private int imagesCount = 0;

	public void mergeFiles()
	{
		try
		{
			List<String> testFileNames = new ArrayList<String>();
			HSSFWorkbook destinationWorkbook = new HSSFWorkbook();

			for (String testFileName : testFileNames)
			{

				File sourceFile = new File(testFileName);

				if (sourceFile.exists())
				{
					System.out.println("\n\nStart executing : " + sourceFile.getAbsolutePath());

					FileInputStream excelSourceTestFile = new FileInputStream(sourceFile);

					HSSFWorkbook sourceWorkbook = new HSSFWorkbook(excelSourceTestFile);
					HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);

					HSSFSheet destinationSheet = destinationWorkbook.createSheet();

					copySheets(destinationSheet, sourceSheet);
					writeToFile(destinationWorkbook, destinationSheet);

				} else
				{
					System.out.println("File doesn't exists: " + testFileName);
				}
			}

		}
		catch (FileNotFoundException e)
		{
			e.printStackTrace();
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	private String destinationFile = "c:\\Merged_File.xls";

	private void writeToFile(HSSFWorkbook workbook, HSSFSheet sheet)
	{
		if (workbook != null && sheet.getPhysicalNumberOfRows() > 0)
		{
			try
			{
				FileOutputStream out = new FileOutputStream(new File(destinationFile));
				workbook.write(out);
				out.close();
				System.out.println(destinationFile + " is written successfully..");

			}
			catch (FileNotFoundException e)
			{
				e.printStackTrace();
			}
			catch (IOException e)
			{
				e.printStackTrace();
			}
		}
	}

	/**
	 * DEFAULT CONSTRUCTOR.
	 */
	public ExcelCopier()
	{
	}

	/**
	 * @param newSheet the sheet to create from the copy.
	 * @param sheet    the sheet to copy.
	 */
	public void copySheets(HSSFSheet newSheet, HSSFSheet sheet)
	{
		copySheets(newSheet, sheet, true);
	}

	/**
	 * @param newSheet  the sheet to create from the copy.
	 * @param sheet     the sheet to copy.
	 * @param copyStyle true copy the style.
	 */
	public void copySheets(HSSFSheet newSheet, HSSFSheet sheet, boolean copyStyle)
	{
		int maxColumnNum = 0;
		Map<Integer, HSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, HSSFCellStyle>() : null;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++)
		{
			HSSFRow srcRow = sheet.getRow(i);
			HSSFRow destRow = newSheet.createRow(i);
			if (srcRow != null)
			{
				// System.out.println("srcRow = " + srcRow.toString());
				copyRow(sheet, newSheet, srcRow, destRow, 0, styleMap);
				if (srcRow.getLastCellNum() > maxColumnNum)
				{
					maxColumnNum = srcRow.getLastCellNum();
				}
			}
		}
		for (int i = 0; i <= maxColumnNum; i++)
		{
			newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
		}
	}

	/**
	 * @param newSheet  the sheet to create from the copy.
	 * @param sheet     the sheet to copy.
	 * @param copyStyle true copy the style.
	 */
	public void addOldSheetToEnd(HSSFSheet newSheet, HSSFSheet sheet, boolean copyStyle)
	{
		int maxColumnNum = 0;
		Map<Integer, HSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, HSSFCellStyle>() : null;
		int lastRowNewSheet = newSheet.getLastRowNum() == 0 ? 0 : newSheet.getLastRowNum() + 1;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++)
		{
			HSSFRow srcRow = sheet.getRow(i);
			HSSFRow destRow = newSheet.createRow(i + lastRowNewSheet);
			if (srcRow != null)
			{
				// System.out.println("srcRow = " + srcRow.toString());
				copyRow(sheet, newSheet, srcRow, destRow, lastRowNewSheet, styleMap);
				if (srcRow.getLastCellNum() > maxColumnNum)
				{
					maxColumnNum = srcRow.getLastCellNum();
				}
			}
		}
		for (int i = 0; i <= maxColumnNum; i++)
		{
			newSheet.setColumnWidth(i + lastRowNewSheet, sheet.getColumnWidth(i));
		}
		List<HSSFPictureData> allPictures = sheet.getWorkbook().getAllPictures();
		int willInsertedImageCount = allPictures.size() - imagesCount;
		ClientAnchor anchor = new HSSFClientAnchor();
		if(willInsertedImageCount == 0)
		{
			return;
		}
		else if(willInsertedImageCount > 1)
		{
			List<HSSFPictureData> picturesToModelRow = new ArrayList<>();
			for (int i = imagesCount; i < allPictures.size(); i++)
			{
				HSSFPictureData hssfPictureData = allPictures.get(i);
				if(allPictures.size() - i == 1)
				{
					insertMainImage(newSheet, hssfPictureData, anchor, lastRowNewSheet);
					break;
				}
				picturesToModelRow.add(hssfPictureData);
			}
			buildModelRow(newSheet, anchor, picturesToModelRow, lastRowNewSheet);
		}
		else
		{
			HSSFPictureData hssfPictureData = allPictures.get(allPictures.size() - 1);
			insertMainImage(newSheet, hssfPictureData, anchor, lastRowNewSheet);
		}
		imagesCount = allPictures.size();

	}

	private void buildModelRow(HSSFSheet newSheet, ClientAnchor anchor, List<HSSFPictureData> picturesToModelRow, int startRow)
	{
		int row1 = startRow;
		int row2 = startRow;
		int col = 12;  //L
		for (int i = 0; i < picturesToModelRow.size(); i++)
		{
			if(i > 0 && i % 3 == 0)
			{
				col = col + 2;
				row1 = startRow;
				row2 = startRow;
			}
			row2 = row2 + 7;
			anchor.setRow1(row1);
			anchor.setCol1(col);
			anchor.setRow2(row2);
			anchor.setCol2(col + 2);
			insertImage(newSheet, picturesToModelRow.get(i), anchor);
			row1 = row2;
		}
	}

	private void insertImage(HSSFSheet newSheet, HSSFPictureData hssfPictureData, ClientAnchor anchor)
	{
		byte[] bytes = hssfPictureData.getData();
		HSSFWorkbook wb = newSheet.getWorkbook();
		int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		Drawing drawing = newSheet.createDrawingPatriarch();
		CreationHelper helper = wb.getCreationHelper();
		ClientAnchor clientAnchor = helper.createClientAnchor();
		clientAnchor.setRow1(anchor.getRow1());
		clientAnchor.setRow2(anchor.getRow2());
		clientAnchor.setCol1(anchor.getCol1());
		clientAnchor.setCol2(anchor.getCol2());
		clientAnchor.setAnchorType(ClientAnchor.DONT_MOVE_AND_RESIZE);
		Picture pict = drawing.createPicture(clientAnchor, pictureIdx);
//		try
//		{
//			String name = "E:\\log\\" + hssfPictureData.getData().length + hssfPictureData.hashCode() + "." + hssfPictureData.suggestFileExtension();
//			File file = new File(name);
//			if(!file.exists())
//			{
//				file.createNewFile();
//			}
//			FileOutputStream fileOutputStream = new FileOutputStream(file);
//			fileOutputStream.write(bytes);
//		}
//		catch (IOException e)
//		{
//			e.printStackTrace();
//		}
//		pict.resize();
	}

	private void insertMainImage(HSSFSheet newSheet, HSSFPictureData hssfPictureData, ClientAnchor anchor, int lastRowNewSheet)
	{
		anchor.setCol1(0);
		anchor.setRow1(lastRowNewSheet);
		anchor.setCol2(5);
		anchor.setRow2(21 + lastRowNewSheet);
		insertImage(newSheet, hssfPictureData, anchor);
	}


	/**
	 * @param srcSheet        the sheet to copy.
	 * @param destSheet       the sheet to create.
	 * @param srcRow          the row to copy.
	 * @param destRow         the row to create.
	 * @param lastRowNewSheet
	 * @param styleMap
	 */
	public void copyRow(HSSFSheet srcSheet, HSSFSheet destSheet, HSSFRow srcRow, HSSFRow destRow,
	                    int lastRowNewSheet, Map<Integer, HSSFCellStyle> styleMap)
	{
		// manage a list of merged zone in order to not insert two times a merged zone
		Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
		destRow.setHeight(srcRow.getHeight());
		// pour chaque row
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++)
		{

			HSSFCell oldCell = srcRow.getCell(j); // old cell
			HSSFCell newCell = destRow.getCell(j); // new cell
			if (oldCell != null)
			{
				if (newCell == null)
				{
					newCell = destRow.createCell(j);
				}
				copyCell(oldCell, newCell, styleMap, lastRowNewSheet);
				// System.out.println("row num: " + srcRow.getRowNum() + " , col: " + (short) oldCell.getColumnIndex());
				try
				{
					CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(), (short) oldCell
							.getColumnIndex(), lastRowNewSheet);
					// System.out.println("mergedRegion: " + mergedRegion);

					if (mergedRegion != null)
					{
//						 System.out.println("Selected merged region: " + mergedRegion.toString());
						CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
								mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
						// System.out.println("New merged region: " + newMergedRegion.toString());
						CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);

						if (isNewMergedRegion(wrapper, mergedRegions))
						{
							// System.out.println("Adding new region");
							mergedRegions.add(wrapper);
							destSheet.addMergedRegion(wrapper.range);
						}
					}
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			}
		}

	}

	/**
	 * @param oldCell
	 * @param newCell
	 * @param styleMap
	 */
	public void copyCell(HSSFCell oldCell, HSSFCell newCell, Map<Integer, HSSFCellStyle> styleMap, int lastRowNewSheet)
	{

		if (styleMap != null)
		{
			if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook())
			{
				newCell.setCellStyle(oldCell.getCellStyle());
			} else
			{
				int stHashCode = oldCell.getCellStyle().hashCode();
				HSSFCellStyle newCellStyle = styleMap.get(stHashCode);
				if (newCellStyle == null)
				{
					newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
					newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
					styleMap.put(stHashCode, newCellStyle);
				}
				newCell.setCellStyle(newCellStyle);
			}
		}
		switch (oldCell.getCellType())
		{
			case HSSFCell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getStringCellValue());
				// System.out.println("oldCell = " + oldCell.getStringCellValue());
				// System.out.println("newCell = " + newCell.getStringCellValue());

				break;
			case HSSFCell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				// System.out.println("oldCell = " + oldCell.getNumericCellValue());
				// System.out.println("newCell = " + newCell.getNumericCellValue());
				break;
			case HSSFCell.CELL_TYPE_BLANK:
				newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
				break;
			case HSSFCell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case HSSFCell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			//TODO: remove hack!!!!
			case HSSFCell.CELL_TYPE_FORMULA:
				StringBuilder builder = new StringBuilder();
				builder.append("H").append(11 + lastRowNewSheet).append("*J").append(11 + lastRowNewSheet).append("*H").append(20 + lastRowNewSheet).append("+");
				builder.append("H").append(12 + lastRowNewSheet).append("*J").append(12 + lastRowNewSheet).append("*H").append(20 + lastRowNewSheet).append("+");
				builder.append("H").append(13 + lastRowNewSheet).append("*J").append(13 + lastRowNewSheet).append("*H").append(20 + lastRowNewSheet);
				newCell.setCellFormula(builder.toString());
				break;
			default:
				break;
		}
	}

	/**
	 * @param sheet   the sheet containing the data.
	 * @param rowNum  the num of the row to copy.
	 * @param cellNum the num of the cell to copy.
	 * @return the CellRangeAddress created.
	 */
	public CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum, short cellNum)
	{
		return getMergedRegion(sheet, rowNum, cellNum, 0);
	}

	public CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum, short cellNum, int startRow)
	{
		for (int i = 0; i < sheet.getNumMergedRegions(); i++)
		{
			CellRangeAddress merged = sheet.getMergedRegion(i);
			int firstRowOldValue = merged.getFirstRow();
			int lastRowOldValue = merged.getLastRow();
			if (startRow != 0)
			{
				merged.setFirstRow(firstRowOldValue + startRow);
				merged.setLastRow(lastRowOldValue + startRow);
//				if(startRow > 21)
//				{
//					System.out.println("Проверяем входит ли ячейка (" + (rowNum + startRow) + "," + cellNum  + ") в границы от ("
//							+ merged.getFirstRow() + "," + merged.getFirstColumn() + ") до ("
//							+  merged.getLastRow() + "," + merged.getLastColumn() + ").");
//				}
			}

			if (merged.isInRange(rowNum + startRow, cellNum))
			{
				CellRangeAddress newMergedAddress = new CellRangeAddress(merged.getFirstRow(), merged.getLastRow(), merged.getFirstColumn(), merged.getLastColumn());
				merged.setFirstRow(firstRowOldValue);
				merged.setLastRow(lastRowOldValue);
				return newMergedAddress;
			}
			merged.setFirstRow(firstRowOldValue);
			merged.setLastRow(lastRowOldValue);
		}
		return null;
	}

	/**
	 * Check that the merged region has been created in the destination sheet.
	 *
	 * @param newMergedRegion the merged region to copy or not in the destination sheet.
	 * @param mergedRegions   the list containing all the merged region.
	 * @return true if the merged region is already in the list or not.
	 */
	private boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
	                                  Set<CellRangeAddressWrapper> mergedRegions)
	{
		return !mergedRegions.contains(newMergedRegion);
	}

	/**
	 * The two cellStyle is the same
	 *
	 * @param cellStyle
	 * @param sourceCellStyle
	 * @param sourceWorkBook
	 * @param destWorkBook
	 * @return
	 */
	private boolean isEqual(HSSFCellStyle cellStyle,
	                        HSSFCellStyle sourceCellStyle, HSSFWorkbook destWorkBook, HSSFWorkbook sourceWorkBook)
	{
		//Judgment as to whether the line style
		if (cellStyle.getWrapText() != sourceCellStyle.getWrapText())
		{
			return false;
		}
		//Alignment is the same
		if (cellStyle.getAlignment() != sourceCellStyle.getAlignment())
		{
			return false;
		}
		if (cellStyle.getVerticalAlignment() != sourceCellStyle.getVerticalAlignment())
		{
			return false;
		}
		//The frame is the same
		if (cellStyle.getBorderBottom() != sourceCellStyle.getBorderBottom())
		{
			return false;
		}
		if (cellStyle.getBorderLeft() != sourceCellStyle.getBorderLeft())
		{
			return false;
		}
		if (cellStyle.getBorderRight() != sourceCellStyle.getBorderRight())
		{
			return false;
		}
		if (cellStyle.getBorderTop() != sourceCellStyle.getBorderTop())
		{
			return false;
		}
		if (cellStyle.getBottomBorderColor() != sourceCellStyle.getBottomBorderColor())
		{
			return false;
		}
		if (cellStyle.getLeftBorderColor() != sourceCellStyle.getLeftBorderColor())
		{
			return false;
		}
		if (cellStyle.getRightBorderColor() != sourceCellStyle.getRightBorderColor())
		{
			return false;
		}
		if (cellStyle.getTopBorderColor() != sourceCellStyle.getTopBorderColor())
		{
			return false;
		}
		//Whether the font
		HSSFFont sourceFont = sourceCellStyle.getFont(sourceWorkBook);
		HSSFFont destFont = cellStyle.getFont(destWorkBook);
		if (destFont.getBoldweight() != sourceFont.getBoldweight())
		{
			return false;
		}
		if (destFont.getCharSet() != sourceFont.getCharSet())
		{
			return false;
		}
		if (destFont.getColor() != sourceFont.getColor())
		{
			return false;
		}
		if (destFont.getFontHeight() != sourceFont.getFontHeight())
		{
			return false;
		}
		if (destFont.getFontHeightInPoints() != sourceFont.getFontHeightInPoints())
		{
			return false;
		}
		if (destFont.getIndex() != sourceFont.getIndex())
		{
			return false;
		}
		if (destFont.getItalic() != sourceFont.getItalic())
		{
			return false;
		}
		if (destFont.getUnderline() != sourceFont.getUnderline())
		{
			return false;
		}
		if (destFont.getStrikeout() != sourceFont.getStrikeout())
		{
			return false;
		}
		if (!destFont.getFontName().equals(sourceFont.getFontName()))
		{
			return false;
		}
		//Other styles are the same
		return true;
	}

//	public void copyImages(HSSFSheet sourceSheet, HSSFSheet targetSheet)
//	{
//		List lst = sourceSheet.getWorkbook().getAllPictures();
//		Hashtable<String, HSSFPicture> picTbl = new Hashtable<String,
//				HSSFPicture>();
//		findShapeInfo(sourceSheet);
//		CreationHelper helper =
//				targetSheet.getWorkbook().getCreationHelper();
//		HSSFPatriarch drawing = null;
//		try
//		{
//			drawing = targetSheet.getDrawingPatriarch();
//		}
//		catch (Exception ex)
//		{
//			ex.printStackTrace();
//		}
//		if (drawing == null)
//			drawing = targetSheet.createDrawingPatriarch();
//
//		for (Shape shape : shapes.values())
//		{
//			if (shape.getPictureFilename() == null)
//				continue;
//
//			HSSFPictureData picData = (HSSFPictureData)
//					lst.get(shape.getPicIndex());
//			shape.setPictureData(picData);
//
//			byte[] data = shape.getPictureData().getData();
//
//			HSSFClientAnchor anchor = (HSSFClientAnchor)
//					helper.createClientAnchor();
//			anchor.setCol1(shape.getCol1());
//			anchor.setCol2(shape.getCol2());
//			anchor.setRow1(shape.getRow1());
//			anchor.setRow2(shape.getRow2());
//			anchor.setDx1(shape.getOffsetCol1());
//			anchor.setDx2(shape.getOffsetCol2());
//			anchor.setDy1(shape.getOffsetRow1());
//			anchor.setDy2(shape.getOffsetRow2());
//
//			// The pictureIdx is wrong
//			int pictureIdx =
//					targetSheet.getWorkbook().addPicture(String.valueOf(anchor.hashCode()),
//							data, shape.getPictureData().getFormat());
//			HSSFPicture pic = drawing.createPicture(anchor,
//					pictureIdx);
//			picTbl.put(String.valueOf(anchor.hashCode()), pic);
//		}
//
//		int index = 1;
//		lst = targetSheet.getWorkbook().getAllPictures();
//		for (int i = 0; i < lst.size(); i++)
//		{
//			index++;
//			HSSFPictureData picData = (HSSFPictureData)
//					lst.get(i);
//
//			if (picData.getCustomId() == null)
//				continue;
//
//			HSSFPicture pic = picTbl.get(picData.getCustomId());
//			if (pic == null)
//				continue;
//
//			pic.setPictureIndex(index);
//		}
//	}
//
//	private void findShapeInfo(HSSFSheet sheet)
//	{
//		try
//		{
//			EscherAggregate escherAggregate = sheet.getDrawingEscherAggregate();
//			if (escherAggregate != null)
//			{
//				EscherContainerRecord escherContainer =
//						escherAggregate.getEscherContainer();
//				iterateContainer(escherContainer, 1);
//			}
//		}
//		catch (Exception ex)
//		{
//			ex.printStackTrace();
//		}
//	}
//
//	private void iterateContainer(EscherContainerRecord escherContainer, int
//			level)
//	{
//		if (escherContainer == null)
//			return;
//
//		List childRecords = escherContainer.getChildRecords();
//		Iterator listIterator = null;
//		org.apache.poi.ddf.EscherSpgrRecord sprgRecord = null;
//		org.apache.poi.ddf.EscherSpRecord spRecord = null;
//		org.apache.poi.ddf.EscherOptRecord optRecord = null;
//		org.apache.poi.ddf.EscherClientAnchorRecord anchrRecord = null;
//		org.apache.poi.ddf.EscherClientDataRecord dataRecord = null;
//		org.apache.poi.ddf.EscherDgRecord dgRecord = null;
//		Object next = null;
//
//		listIterator = childRecords.iterator();
//		while (listIterator.hasNext())
//		{
//			next = listIterator.next();
//
//			// logger.debug("next: " + next);
//
//			if (next instanceof org.apache.poi.ddf.EscherContainerRecord)
//				iterateContainer((org.apache.poi.ddf.EscherContainerRecord) next,
//						++level);
//			else if (next instanceof org.apache.poi.ddf.EscherSpgrRecord)
//				sprgRecord = (org.apache.poi.ddf.EscherSpgrRecord) next;
//			else if (next instanceof org.apache.poi.ddf.EscherSpRecord)
//				spRecord = (org.apache.poi.ddf.EscherSpRecord) next;
//			else if (next instanceof org.apache.poi.ddf.EscherOptRecord)
//			{
//				optRecord = (org.apache.poi.ddf.EscherOptRecord) next;
//				String key = String.valueOf(level);
//				if (shapes.containsKey(key))
//					shapes.get(key).setOptRecord(optRecord);
//				else
//				{
//					ShapeInfo shape = new ShapeInfo(level);
//					shape.setOptRecord(optRecord);
//					shapes.put(key, shape);
//				}
//			}
//			else if (next instanceof org.apache.poi.ddf.EscherClientAnchorRecord)
//			{
//				anchrRecord = (org.apache.poi.ddf.EscherClientAnchorRecord) next;
//				String key = String.valueOf(level);
//				if (shapes.containsKey(key))
//					shapes.get(key).setAnchorRecord(anchrRecord);
//				else
//				{
//					ShapeInfo shape = new ShapeInfo(level);
//					shape.setAnchorRecord(anchrRecord);
//					shapes.put(key, shape);
//				}
//			}
//			else if (next instanceof org.apache.poi.ddf.EscherClientDataRecord)
//				dataRecord = (org.apache.poi.ddf.EscherClientDataRecord) next;
//			else if (next instanceof org.apache.poi.ddf.EscherDgRecord)
//				dgRecord = (org.apache.poi.ddf.EscherDgRecord) next;
//		}
//	}
}