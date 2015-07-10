package by.dabl.model.copier;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.*;

import by.dabl.model.CellRangeAddressWrapper;

/**
 *  Implementation for working with old .xsl files, Excel 1997-2003 versions.
 */

public class HSSFExcelCopier implements ExcelCopier
{

	/**
	 * Stateful variable. Contains already copied images count.
	 * Maybe will be removed in future, after change system design.
	 */

	private int alreadyCopiedImagesCount = 0;

	protected void mergeFiles()
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
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	private String destinationFile = "c:\\Merged_File.xls";

	protected void writeToFile(HSSFWorkbook workbook, HSSFSheet sheet)
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
	public HSSFExcelCopier()
	{
	}

	@Override
	public void copySheets(Sheet newSheet, Sheet sheet)
	{
		copySheets(newSheet, sheet, true);
	}

	@Override
	public void copySheets(Sheet newSheet, Sheet sheet, boolean copyStyle)
	{
		int maxColumnNum = 0;
		Map<Integer, CellStyle> styleMap = (copyStyle) ? new HashMap<Integer, CellStyle>() : null;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++)
		{
			Row srcRow = sheet.getRow(i);
			Row destRow = newSheet.createRow(i);
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
	public void addOldSheetToEnd(Sheet newSheet, Sheet sheet, boolean copyStyle)
	{
		int maxColumnNum = 0;
		Map<Integer, CellStyle> styleMap = (copyStyle) ? new HashMap<Integer, CellStyle>() : null;
		int lastRowNewSheet = newSheet.getLastRowNum() == 0 ? 0 : newSheet.getLastRowNum() + 1;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++)
		{
			Row srcRow = sheet.getRow(i);
			Row destRow = newSheet.createRow(i + lastRowNewSheet);
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
		List<? extends PictureData> allPictures = sheet.getWorkbook().getAllPictures();
		int willInsertedImageCount = allPictures.size() - alreadyCopiedImagesCount;
		//TODO: think about abstraction
		ClientAnchor anchor = new HSSFClientAnchor();
		if(willInsertedImageCount == 0)
		{
			return;
		}
		else if(willInsertedImageCount > 1)
		{
			List<PictureData> picturesToModelRow = new ArrayList<>();
			for (int i = alreadyCopiedImagesCount; i < allPictures.size(); i++)
			{
				PictureData hssfPictureData = allPictures.get(i);
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
			PictureData hssfPictureData = allPictures.get(allPictures.size() - 1);
			insertMainImage(newSheet, hssfPictureData, anchor, lastRowNewSheet);
		}
		alreadyCopiedImagesCount = allPictures.size();

	}


	//TODO: remove hardcode
	private void buildModelRow(Sheet newSheet, ClientAnchor anchor, List<PictureData> picturesToModelRow, int startRow)
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

	private void insertImage(Sheet newSheet, PictureData pictureData, ClientAnchor anchor)
	{
		byte[] bytes = pictureData.getData();
		Workbook wb = newSheet.getWorkbook();
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
	}

	private void insertMainImage(Sheet newSheet, PictureData pictureData, ClientAnchor anchor, int lastRowNewSheet)
	{
		//TODO: remove hardcode
		anchor.setCol1(0);
		anchor.setRow1(lastRowNewSheet);
		anchor.setCol2(5);
		anchor.setRow2(21 + lastRowNewSheet);
		insertImage(newSheet, pictureData, anchor);
	}


	public void copyRow(Sheet srcSheet, Sheet destSheet, Row srcRow, Row destRow,
	                    int lastRowNewSheet, Map<Integer, CellStyle> styleMap)
	{
		// manage a list of merged zone in order to not insert two times a merged zone
		Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
		destRow.setHeight(srcRow.getHeight());
		// pour chaque row
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++)
		{

			Cell oldCell = srcRow.getCell(j); // old cell
			Cell newCell = destRow.getCell(j); // new cell
			if (oldCell != null)
			{
				if (newCell == null)
				{
					newCell = destRow.createCell(j);
				}
				copyCell(oldCell, newCell, styleMap, lastRowNewSheet);
				// System.out.println("row num: " + srcRow.getRowNum() + " , col: " + (short) oldCell.getColumnIndex());
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
		}
	}

	@Override
	public void copyCell(Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap, int lastRowNewSheet)
	{
		if (styleMap != null)
		{
			if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook())
			{
				newCell.setCellStyle(oldCell.getCellStyle());
			} else
			{
				int stHashCode = oldCell.getCellStyle().hashCode();
				CellStyle newCellStyle = styleMap.get(stHashCode);
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
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getStringCellValue());
				// System.out.println("oldCell = " + oldCell.getStringCellValue());
				// System.out.println("newCell = " + newCell.getStringCellValue());

				break;
			case Cell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				// System.out.println("oldCell = " + oldCell.getNumericCellValue());
				// System.out.println("newCell = " + newCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_BLANK:
				newCell.setCellType(Cell.CELL_TYPE_BLANK);
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			//TODO: remove hardcode!!!!
			case Cell.CELL_TYPE_FORMULA:
				String oldCellFormula = oldCell.getCellFormula();
				if(oldCellFormula.startsWith("J11*H20"))
				{
					StringBuilder builder = new StringBuilder();
					for (int i = 11; i < 19; i++)
					{
						builder.append("H").append(i + lastRowNewSheet).append("*J").append(i + lastRowNewSheet).append("*H").append(20 + lastRowNewSheet).append("+");
					}
					builder.append("H").append(19 + lastRowNewSheet).append("*J").append(19 + lastRowNewSheet).append("*H").append(20 + lastRowNewSheet);
					oldCellFormula = builder.toString();
				}
				if(oldCellFormula.equals("BB20"))
				{
					oldCellFormula = oldCellFormula.replace("20", String.valueOf(20 + lastRowNewSheet));
				}
				newCell.setCellFormula(oldCellFormula);
				break;
			default:
				break;
		}
	}

	@Override
	public CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum)
	{
		return getMergedRegion(sheet, rowNum, cellNum, 0);
	}

	public CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum, int startRow)
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
}
