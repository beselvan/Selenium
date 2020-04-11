package driverScript;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Properties;
import java.util.StringTokenizer;
import org.apache.log4j.xml.DOMConfigurator;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.Test;
import excelReader.ReadExcel;
import excelReader.UpdateResult;
import objectRepository.LoadPropertySingleton;
import objectRepository.ReadObject;
import operations.UIOperations;

public class ExecuteTest 
{
	static final Logger loggerET = LogManager.getLogger(ExecuteTest.class.getName());
	static LoadPropertySingleton objectLoad = LoadPropertySingleton.getInstance();
	public static WebDriver driver = null;
	public boolean checkFlag = false;
	public boolean prevTC = false;
	public static boolean driverFlag = true;
	public int 					tcNameRow;
	public String passCase = LoadPropertySingleton.configResourceBundle.getString("TESTCASE_PASS");
	public String failCase = LoadPropertySingleton.configResourceBundle.getString("TESTCASE_FAIL");
	public ArrayList<String> t_Status = new ArrayList<String>();
	public String filePath_TC = LoadPropertySingleton.configResourceBundle.getString("TestCaseFilePath");
	public String fileName_TC = LoadPropertySingleton.configResourceBundle.getString("TestCaseFileName");
	public String filePath_R = LoadPropertySingleton.configResourceBundle.getString("ReportFilePath");
	public String fileName_R = LoadPropertySingleton.configResourceBundle.getString("ReportFileName");
	public String evidencePath = LoadPropertySingleton.configResourceBundle.getString("EvidenceFilePath");
	public String webDriverLocation = LoadPropertySingleton.configResourceBundle.getString("WebDriverPath");
	public String browserName = LoadPropertySingleton.configResourceBundle.getString("Browser");
	public String geckoDriverPath = LoadPropertySingleton.configResourceBundle.getString("geckoDriver");
	public static boolean loginCheckFlag = true;
	
	@Test
	public void execute(String taskcode) throws Exception 
	{
		loggerET.info("ExecuteTest.java : execute()"+taskcode);
		initializeRunTimeTask(browserName,driverFlag,webDriverLocation,geckoDriverPath);
		ReadExcel objReadExcel = new ReadExcel();
		ReadObject objReadObject = new ReadObject();
		Properties propertiesObj = objReadObject.getObjectRepository();
		UIOperations objUIOperations = new UIOperations(driver);
		UpdateResult objUpdateResult = new UpdateResult();
				
		Sheet sheet = objReadExcel.readSheetContents(filePath_TC, taskcode+"_"+fileName_TC , taskcode);
		int rowCount = sheet.getLastRowNum()- sheet.getFirstRowNum();
		int totRowCount = sheet.getLastRowNum();
		System.out.println("Test Case File Name ==>: " + taskcode+"_"+fileName_TC +"Test Case Sheet Name==>: "+taskcode);
		System.out.println("Last Row Number: " + totRowCount + "   Row Count: " + rowCount);
		tcNameRow = 1;
		FileInputStream objFileInputStream = objUpdateResult.getFileStreamObject(filePath_TC, taskcode+"_"+fileName_TC );
		XSSFWorkbook objXSSFWorkbook =objUpdateResult.getWorkBookObject(objFileInputStream);
		Sheet actSheet =objUpdateResult.getSheetObject(objXSSFWorkbook,taskcode);
		//workbook object and sheet object should be retrieved only once
		//Create a loop over all the rows of excel file to read it
		for (int i = 1; i <= totRowCount; i++) 
		{
			loggerET.info("execute() Test case Processing Starts with ==> " + i);
			loggerET.info("I = " + i);
			UIOperations.alertMessage = "";
			Row row = sheet.getRow(i);
			try
			{
/*************************************** To Update the Test Case Level Status PASS or FAIL - Starts************************************************************************************/				
				if(prevTC=true && row.getCell(0).toString().length()!=0)
				{
					loggerET.info("TC Name: " + row.getCell(0).toString() + "---> " + "Previous TestCase Flag: " + prevTC);
					
					if(t_Status.contains("False") && (checkFlag))
					{
						objUpdateResult.writeOutputWorkBook(objXSSFWorkbook,actSheet,taskcode, tcNameRow, failCase,"");
						loggerET.info("1_Excel Updated in row no: " + tcNameRow + " FAIL");
						loggerET.info("Updated Test Case-Result in the Report as FAIL when atleast one step in the test case fails ");
						t_Status = new ArrayList<String>();
					}
					else if(!t_Status.contains("False") && (checkFlag))
					{
						objUpdateResult.writeOutputWorkBook(objXSSFWorkbook,actSheet,taskcode, tcNameRow, passCase,"");
						loggerET.info("2_Excel Updated in row no: " + tcNameRow + " PASS");
						loggerET.info("Updated Test Case-Result in the Report as PASS when all the test steps are Passed ");
						t_Status = new ArrayList<String>();
					}
					prevTC = false;
				}
/***************************************To Update the Test Case Status PASS or FAIL - Ends************************************************************************************/				
			}
			catch(Exception ex)
			{
				loggerET.error("Exception Caught :ex : "+ex);
				ex.printStackTrace();
			}
			try
			{
			if(row.getCell(0).toString().length()!=0)
			{
				if(loginCheckFlag == false)		//	LoginCheckFlag is to skip the successive test steps when the login itself failed and update the test case level status as FAIL.
				{
					objUpdateResult.writeOutputWorkBook(objXSSFWorkbook,actSheet,taskcode, tcNameRow, failCase,"");
					loggerET.info("Test Case Result updated as FAIL - LoginCheckFlag - False");
					loginCheckFlag = true;	//	LoginCheckFlag changed to TRUE so that the next test case can be executed.
				}
				if(row.getCell(1).toString().equalsIgnoreCase(LoadPropertySingleton.configResourceBundle.getString("YES")))
                {
                	//Print the new test case name when it gets started
					loggerET.info("New Testcase-> "+row.getCell(0).toString() + " Run_Flag: " + row.getCell(1).toString() + " Started");
					if(objUIOperations.Pass_SnapShot.equalsIgnoreCase("TRUE") || objUIOperations.Fail_SnapShot.equalsIgnoreCase("TRUE"))
					{
						objUIOperations.tcName(row.getCell(0).toString());
					}
 	               	checkFlag = true;		//	To make the test steps to be executed for the Execution Flag - "Yes" test cases, keep checkFlag as True
                	prevTC = true;			//	To update the status of the Test Case Level, keep previousTestCase as True
                	tcNameRow = i;			//	To know the row of the Test Case Name which has Execution Flag as "Yes" so that TestCaseLevel status can be updated
                }
                else
                {
                	checkFlag = false;
                	loggerET.info("New Testcase-> "+row.getCell(0).toString() + " Run_Flag: " + row.getCell(1).toString());
                }
     
			}
			else
			{
				loggerET.info("Login Check Flag: " + loginCheckFlag);
				if(checkFlag && loginCheckFlag)		//	If only the ExecutionCheckFlag and LoginCheckFlag is true the test steps will be executed
				{
					//Print test step detail in console
					loggerET.info(row.getCell(2).toString()+"----"+ row.getCell(3).toString()+"----"+
							row.getCell(4).toString()+"----"+ row.getCell(5).toString());
					loggerET.info("Sheet Name: " + taskcode);
					
					while(objUIOperations.isAlert()==true)
		            {
		            	if(objUIOperations.alertFlag)
		            	{
		            		objUIOperations.alertMessage = objUIOperations.getCancelAlert();
		            	}
		            	else
		            	{
		            		objUIOperations.alertMessage = objUIOperations.getAcceptAlert();
		            		if(objUIOperations.loginAlertCheck)		//	If "Invalid User" Alert is thrown, loginCheckFlag is changed to FALSE
		            		{
		            			loginCheckFlag = false;
		            			loggerET.info("Invalid User - 1");
		            			break;
		            		}
		            	}
		            }
								
					if(!objUIOperations.alertMessage.equalsIgnoreCase(""))
					{
					       if(objUIOperations.alertMessage.contains("Invalid") || objUIOperations.alertMessage.contains("Not") || objUIOperations.alertMessage.contains("invalid") || objUIOperations.alertMessage.contains("not") || objUIOperations.alertMessage.contains("Rejected") || objUIOperations.alertMessage.contains("rejected"))
					       {
						   			objUpdateResult.writeOutputWorkBook(objXSSFWorkbook, actSheet, taskcode, i, "F",UIOperations.alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
						     		System.out.println("Alert Message - Updated Result - Fail due to INVALID or NOT");
						     		t_Status.add("False");
						     		
						   }
						   else
						   {
						           objUpdateResult.writeOutputWorkBook(objXSSFWorkbook, actSheet, taskcode, i, "P",UIOperations.alertMessage);
						           System.out.println("Alert Message - Updated Result as Pass");
						           t_Status.add("True");
						           
						   }
						        
					}
					else
					{
								objUpdateResult.writeOutputWorkBook(objXSSFWorkbook, actSheet, taskcode, i, "P",UIOperations.alertMessage);
								System.out.println("No Alert Message - Pass");
						   		t_Status.add("True");
					}
				if(row.getCell(6).getCellType()==row.getCell(6).CELL_TYPE_NUMERIC)
				{
					loggerET.info("Test Data Value: " + NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()));
					loggerET.info("Calling objUIOperations when cell type is NUMERIC ");
					//Call perform function to perform objUIOperations on UI when cell type(Test Data) is Numeric 
					objUIOperations.perform(propertiesObj, row.getCell(0).toString(), row.getCell(2).toString(), row.getCell(3).toString(),
							row.getCell(4).toString(), row.getCell(5).toString(), NumberToTextConverter.toText(row.getCell(6).getNumericCellValue()), row.getCell(7).toString(), row.getCell(8).CELL_TYPE_NUMERIC, i, tcNameRow,t_Status, webDriverLocation,objXSSFWorkbook,actSheet, taskcode );
				}
				else
				{
					loggerET.info("Calling Operation when cell type is NON-NUMERIC ");
					//Call perform function to perform operation on UI when cell type(Test Data) is Non-Numeric 
					objUIOperations.perform(propertiesObj, row.getCell(0).toString(), row.getCell(2).toString(), row.getCell(3).toString(),
							row.getCell(4).toString(), row.getCell(5).toString(), row.getCell(6).toString(), row.getCell(7).toString(), row.getCell(8).CELL_TYPE_NUMERIC, i, tcNameRow,t_Status, webDriverLocation,objXSSFWorkbook,actSheet, taskcode);
				}
			}
		}
			}
			catch (Exception exp)
			{
				loggerET.info("Catch block  exp: " + exp);
				loggerET.error("Exception  Caught exp ==>"+exp);
				objUpdateResult.writeOutputWorkBook(objXSSFWorkbook,actSheet, taskcode, tcNameRow, LoadPropertySingleton.configResourceBundle.getString("TESTCASE_FAIL"),"");
				exp.printStackTrace();
			}
			loggerET.info("Array List: " + t_Status);
	   }
		objUpdateResult.closeWorkBook(objXSSFWorkbook,filePath_R, fileName_R,objFileInputStream);
	}
	public static void killTask() throws Exception 
	{
		Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
		Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
		
	}
   	public void initializeRunTimeTask(String browserName,Boolean driverFlag,String webDriverLocation,String geckoDriverPath) 
    		throws Exception
    {
    	DOMConfigurator.configure("Logs.xml");
		if(driverFlag)
		{
			Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
			Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
			Runtime.getRuntime().exec("taskkill /F /IM firefox.exe");
			Runtime.getRuntime().exec("taskkill /F /IM geckodriver.exe");
			loggerET.info("Closing the Existing IE Browsers");
		}
		if(browserName.equalsIgnoreCase("IE"))
		{
/******************************* Launching IE Browser ****************************************************************/			
			System.setProperty("webdriver.ie.driver", webDriverLocation);  
			driver = new InternetExplorerDriver();
			loggerET.info("IE Driver Launched Successfully");
		}
		else if (browserName.equalsIgnoreCase("FireFox"))
		{
/******************************* Launching Firefox Browser ****************************************************************/	
		System.setProperty("webdriver.gecko.driver",geckoDriverPath);
		loggerET.info("Firefox Driver Launched Successfully");
		}
    	
    }
/************************************* Key Word Driven Ends *********************************************************/
	public static void main(String args[]) throws Exception 
	{
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	    Date now = new Date();
	    String startTime=sdf.format(now);
	    System.out.println("Start time ==> "+startTime);
		ExecuteTest objExecuteTest = new ExecuteTest();
		String taskCode = LoadPropertySingleton.configResourceBundle.getString("ListOfTaskCodes");
		StringTokenizer token = new StringTokenizer(taskCode, ",");
		while(token.hasMoreTokens()){
			String taskCodeName = token.nextToken();
			System.out.println("Task Code: " + taskCodeName);
			objExecuteTest.execute(taskCodeName);
			driverFlag = false;
			System.out.println("Driver Flag 2: " + driverFlag);
		}
		killTask();
		Date end = new Date();
		String endTime=sdf.format(end);
		System.out.println("startTime ==> "+startTime +"End time ==> "+endTime);
	}
}
