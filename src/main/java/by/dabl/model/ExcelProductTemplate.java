package by.dabl.model;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;


/**
 * Created by Nikita Tkachuk
 */
public class ExcelProductTemplate
{
	public static final ExcelProductTemplate INSTANCE = new ExcelProductTemplate();

	private ExcelProductTemplate()
	{
	}

	public Sheet getTemplateSheet()
	{
		try(InputStream stream = getClass().getClassLoader().getResourceAsStream("template.xls"))
		{
			Workbook workbook = new HSSFWorkbook(stream);
			return workbook.getSheetAt(0);
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return null;
	}



}
