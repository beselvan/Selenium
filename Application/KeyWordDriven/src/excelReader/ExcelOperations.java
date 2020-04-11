package excelReader;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.Calendar;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;

public class ExcelOperations 
{
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null;
	private HSSFRow row = null;
	private HSSFCell cell = null;
	public static int pat = 1;
	public static int pos = 0;
	public static Drawing drawing1;

	public void Xls_Reader(String path)
	{

		this.path = path;
		try
		{
			fis = new FileInputStream(path);

			workbook = new HSSFWorkbook(fis);

			fis.close();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}

	}
	
	public String getCellData(String sheetName, String colName, int rowNum,	int headingRowNo)
	{
		try
		{
			if (rowNum <= 0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(headingRowNo);
			for (int i = 0; i < row.getLastCellNum(); i++)
			{
				if (!(row.getCell(i) == null))
				{
					if (row.getCell(i).getStringCellValue().trim()
							.equals(colName.trim()))
					{
						col_Num = i;
						break;
					}
				}

			}
			if (col_Num == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(headingRowNo + rowNum);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);

			if (cell == null)
				return "";
			
			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == 0)
			{
				String value = BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString();
				if (value.contains("."))
				{
					value = value.substring(0, value.indexOf("."));
					return value;
				}
				else
				{
					return BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString();
				}
			}
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
					|| cell.getCellType() == Cell.CELL_TYPE_FORMULA)
			{

				String cellText = String.valueOf(cell.getNumericCellValue());

				if (HSSFDateUtil.isCellDateFormatted(cell))
				{
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR)))
							.substring(2);
					cellText = cal.get(Calendar.DAY_OF_MONTH) + "/"
							+ cal.get(Calendar.MONTH) + 1 + "/" + cellText;

				}

				return cellText;
			}
			else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());

		}
		catch (Exception e)
		{

			e.printStackTrace();
			return "row " + rowNum + " or column " + colName
					+ " does not exist in xls";
		}
	}
	
	
	public String getCellData(String sheetName, int colNum, int rowNum)
	{
		try
		{
			// if(rowNum <=0)
			// return "";

			int index = workbook.getSheetIndex(sheetName);

			if (index == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum);
			if (row == null)
				return "";
			cell = row.getCell(colNum);
			if (cell == null)
				return "";
			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
					|| cell.getCellType() == Cell.CELL_TYPE_FORMULA)
			{

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell))
				{
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR)))
							.substring(2);
					cellText = cal.get(Calendar.MONTH) + 1 + "/"
							+ cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;
				}
				return cellText;
			}
			else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		}
		catch (Exception e)
		{

			e.printStackTrace();
			return "row " + rowNum + " or column " + colNum
					+ " does not exist  in xls";
		}
	}
	
	
	public String getCellDataValue(String sheetName, String colName, int rowCnt)
	{

		int col_Num = -1;
		int rowNum = 0;
		int lastRowNum;
		int lastCellNum;
		boolean targetRow = false;
		try
		{

			int index = workbook.getSheetIndex(sheetName);

			if (index == -1)
				return null;

			sheet = workbook.getSheetAt(index);

			lastRowNum = sheet.getLastRowNum();

			for (int i = 0; i < lastRowNum; i++)
			{

				row = sheet.getRow(i);

				if (row == null)
					continue;
				lastCellNum = row.getLastCellNum();

				for (int j = 0; j < lastCellNum; j++)
				{
					// System.out.println(row.getCell(i).getStringCellValue().trim());
					if (row.getCell(j).getStringCellValue().trim()
							.equals(colName.trim()))
					{
						col_Num = j;
						row = sheet.getRow(i + rowCnt);
						targetRow = true;
						break;
					}

				}
				if (targetRow)
					break;
			}

			if (col_Num == -1)
				return null;

			cell = row.getCell(col_Num);

			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
					|| cell.getCellType() == Cell.CELL_TYPE_FORMULA)
			{

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell))
				{
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR)))
							.substring(2);
					cellText = cal.get(Calendar.MONTH) + 1 + "/"
							+ cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;
				}
				return cellText;
			}
			else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		}
		catch (Exception e)
		{

			e.printStackTrace();
			return "row " + rowNum + " or column " + col_Num
					+ " does not exist  in xls";
		}

	}
	
	
	public boolean setCellData(String sheetName, String colName, int rowNum, String data)
	{
		try
		{
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			if (rowNum <= 0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			
			for (int i = 0; i < row.getLastCellNum(); i++)
			{
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;

			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			
			cell.setCellValue(data);
			fileOut = new FileOutputStream(path);

			workbook.write(fileOut);

			fileOut.close();

		}
		catch (Exception e)
		{
			e.printStackTrace();
			return false;
		}
		return true;
	}
	
	public boolean setCellData(String sheetName, String colName, int rowNum, String data, String url)
	{
		// System.out.println("setCellData setCellData******************");
		try
		{
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);

			if (rowNum <= 0)
				return false;

			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;

			sheet = workbook.getSheetAt(index);
			// System.out.println("A");
			row = sheet.getRow(0);
			
			colNum = 0;
			if (colNum == -1)
				return false;
			sheet.autoSizeColumn(colNum); //
			row = sheet.getRow(rowNum);
			if (row == null)
				row = sheet.createRow(rowNum);

			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			cell.setCellValue(data);
			CreationHelper createHelper = workbook.getCreationHelper();

			// cell style for hyperlinks
			// by default hypelrinks are blue and underlined
			CellStyle hlink_style = workbook.createCellStyle();
			HSSFFont hlink_font = workbook.createFont();
			hlink_font.setUnderline(HSSFFont.U_SINGLE);
			hlink_font.setColor(IndexedColors.BLUE.getIndex());
			hlink_style.setFont(hlink_font);
			// hlink_style.setWrapText(true);

			Hyperlink link = createHelper
					.createHyperlink(HSSFHyperlink.LINK_FILE);

			link.setAddress(url);

			cell.setHyperlink(link);
			cell.setCellStyle(hlink_style);

			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);

			fileOut.close();

		}
		catch (Exception e)
		{
			e.printStackTrace();
			return false;
		}
		return true;
	}
}
