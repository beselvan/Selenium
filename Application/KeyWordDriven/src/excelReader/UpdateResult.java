package excelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateResult 
{
	public void writeOutputWorkBook(XSSFWorkbook xssfWorkbook,Sheet targetSheet ,String sheetName,int rowNo, String status, String errorMsg) throws IOException
	{
		XSSFRow r = (XSSFRow) targetSheet.getRow(rowNo);
		XSSFCell c = r.createCell(9);
		CellStyle style = xssfWorkbook.createCellStyle();
		Font font = xssfWorkbook.createFont();
		//Add by Sreenu for Error Message update
		XSSFCell column = r.createCell(10);
		CellStyle newStyle = xssfWorkbook.createCellStyle();
		Font statusFont = xssfWorkbook.createFont();
		
		if(status.equalsIgnoreCase("P"))
		{
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setFontHeightInPoints((short) 9);
		}
		else if(status.equalsIgnoreCase("PASS"))
		{
			style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			font.setFontHeightInPoints((short) 18);
		}
		else if(status.equalsIgnoreCase("F"))
		{
			font.setColor(IndexedColors.RED.getIndex());
			font.setFontHeightInPoints((short) 9);
			//Add by Sreenu for Error Message update
			statusFont.setColor(IndexedColors.BLACK.getIndex());
			newStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
			newStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			statusFont.setFontHeightInPoints((short) 10);
		}
		else
		{
			style.setFillForegroundColor(IndexedColors.RED.getIndex());
			font.setFontHeightInPoints((short) 18);
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		style.setFont(font);
		c.setCellStyle(style);
		c.setCellValue(status);
		//Add by Sreenu for Error Message update
		newStyle.setFont(statusFont);
		column.setCellStyle(newStyle);
		column.setCellValue(errorMsg);
	}
	
	public FileInputStream getFileStreamObject(String filePath,String fileName) throws IOException
	{
		File targetFile = new File(filePath + fileName);
		FileInputStream fileInputStream = new FileInputStream(targetFile);
		return fileInputStream;
		
	}
	public XSSFWorkbook getWorkBookObject(FileInputStream fileInputStream) throws IOException
	{
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
		return xssfWorkbook;
		
	}
	public XSSFSheet getSheetObject(XSSFWorkbook xssfWorkbook,String sheetName) throws IOException
	{
		XSSFSheet xssfSheet = xssfWorkbook.getSheet(sheetName);
		return xssfSheet;
	}
	
	public void closeWorkBook(XSSFWorkbook xssfWorkbook,String filePath,String fileName,FileInputStream fileInputStream) throws IOException
	{
		System.out.println("closeWorkBook   xssfWorkbook: " + xssfWorkbook +"filePath ==>"+filePath +"fileName"+fileName);
		File targetFile = new File(filePath + fileName);
		System.out.println("closeWorkBook   targetFile: " + targetFile);
		FileOutputStream fileOutputStream = new FileOutputStream(targetFile);
		System.out.println("closeWorkBook   fileOutputStream: " + fileOutputStream);
		xssfWorkbook.write(fileOutputStream);
		System.out.println("After write workbook");
		fileInputStream.close();
		fileOutputStream.close();
		xssfWorkbook.close();
	}
}
