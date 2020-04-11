package excelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
{
	public Sheet readSheetContents(String filePath,String fileName,String sheetName) throws IOException
	{
		File objFile =    new File(filePath+"\\"+fileName);
		FileInputStream objFileInputStream = new FileInputStream(objFile);
	    Workbook objWorkbook = null;
	    Sheet sheet =null;
	    String fileExtensionName = fileName.substring(fileName.indexOf("."));
	    try
    	{
	    	if(fileExtensionName.equals(".xlsx"))
	    	  		objWorkbook = new XSSFWorkbook(objFileInputStream);
	        else if(fileExtensionName.equals(".xls"))
	        		objWorkbook = new HSSFWorkbook(objFileInputStream);
	    	if (objWorkbook!=null)
	    		sheet = objWorkbook.getSheet(sheetName);
	   	}	
	    catch(Exception exp) 
	   	{
			System.out.println("Exception Message in ReadExcel.java : readSheetContents()");
			exp.printStackTrace();
	    }
	   finally
	    {
	    	if (objFileInputStream!=null)
	    		objFileInputStream.close();
	    	if (objWorkbook!=null)
	    		objWorkbook.close();
	    }
	    return sheet;    
	}
}
