package by.dabl.model.copier;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Created with IntelliJ IDEA.
 * User: nikita tkachuk
 * Date: 7/10/2015
 * Time: 10:20 AM
 */
public interface ExcelCopier
{
	/**
	 * @param newSheet the sheet to create from the copy.
	 * @param sheet    the sheet to copy.
	 */

	void copySheets(Sheet newSheet, Sheet sheet);


	/**
	 * @param newSheet  the sheet to create from the copy.
	 * @param sheet     the sheet to copy.
	 * @param copyStyle true copy the style.
	 */

	void copySheets(Sheet newSheet, Sheet sheet, boolean copyStyle);


	/**
	 * @param srcSheet        the sheet to copy.
	 * @param destSheet       the sheet to create.
	 * @param srcRow          the row to copy.
	 * @param destRow         the row to create.
	 * @param lastRowNewSheet
	 * @param styleMap
	 */

	void copyRow(Sheet srcSheet, Sheet destSheet, Row srcRow, Row destRow,
						int lastRowNewSheet, Map<Integer, CellStyle> styleMap);

	/**
	 * @param oldCell
	 * @param newCell
	 * @param styleMap
	 */

	void copyCell(Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap, int lastRowNewSheet);

	/**
	 * @param sheet   the sheet containing the data.
	 * @param rowNum  the num of the row to copy.
	 * @param cellNum the num of the cell to copy.
	 * @return the CellRangeAddress created.
	 */

	CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum);

	/**
	 * @param sheet   the sheet containing the data.
	 * @param rowNum  the num of the row to copy.
	 * @param cellNum the num of the cell to copy.
	 * @param startRow the number of last row. Using for adding merge region in end of sheet.
	 * @return the CellRangeAddress created.
	 */

	CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum, int startRow);
}
