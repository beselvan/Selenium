package operations;

import java.awt.AWTException;
import java.awt.HeadlessException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import javax.imageio.ImageIO;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.log4j.xml.DOMConfigurator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchWindowException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import driverScript.ExecuteTest;
import excelReader.UpdateResult;
import objectRepository.LoadPropertySingleton;

public class UIOperations 
{
	public static WebDriver driver = null;
	private String fileName_E;
	public static WebElement element;
	public static JavascriptExecutor js;
	public static String imagedir = "TestEvidenceImage";
    public static int imagecounter = 1;
    static DateFormat formatter = new SimpleDateFormat("yyyy/MM/dd");
    ExecuteTest objExecuteTest = new ExecuteTest();
    UpdateResult objUpdateResult = new UpdateResult();
    public static boolean alertFlag = false;
    public static boolean loginAlertCheck = false;
    public static String alertMessage = "";
    public static String Pass_SnapShot = LoadPropertySingleton.configResourceBundle.getString("Pass_ScreenShot");
	public static String Fail_SnapShot = LoadPropertySingleton.configResourceBundle.getString("Fail_ScreenShot");
    static final Logger loggerUI = LogManager.getLogger(UIOperations.class.getName());
    String tcStatus = LoadPropertySingleton.configResourceBundle.getString("TESTCASE_PASS");
    public UIOperations(WebDriver driver)
    {
        UIOperations.driver = driver;
    }
   
    public void perform(Properties objProperties,String testCase, String operation,String objectName,String objectType,
    	String frame, String value, String delay, int slNo, int row, int tcNameRow, ArrayList<String> t_Status, 
    	String webdriveLocation,XSSFWorkbook xssfWorkbook,Sheet actSheet,String taskCode) throws Exception
    {
        long i = (long) Double.parseDouble(delay);
        loginAlertCheck = false;		//	The LoginAlertCheck is changed to FALSE so that the status of the next CLICK BUTTON will be updated properly
        WebDriverWait wait = new WebDriverWait(driver,i);
        
        while(isAlert()==true)
		{
			if(alertFlag)
			{
				alertMessage = getCancelAlert();
				loggerUI.info("Alert - Cancel Clicked ");
			}
			else
			{
				alertMessage = getAcceptAlert();
				loggerUI.info("Alert - OK Clicked ");
			}
		}
        if(!frame.equals(""))
        {
        	waitForFrame(frame,objectType);
        }
        System.out.println("Operation Name: " + operation);
        Actions act = new Actions(driver);
        switch (operation.toUpperCase()) 
        {
          case "CLICKBUTTON":
            //	Perform click on Button
        	  loggerUI.info("CLICK BUTTON_SLEEP: " + i);
        	  loggerUI.info("Object Name inside Click Button: " + objProperties.getProperty(objectName));
        	if(objectName.equals("authorization.ok"))
        	{
        		loggerUI.info("Before WebDriver Wait: ");
        		WebDriverWait wait1 = new WebDriverWait(driver,12);
        		try
        		{
        			wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(objProperties.getProperty(objectName))));
        			t_Status.add("True");
        		}
        		catch(Exception e)
        		{
					loggerUI.info(objectName + "AUTHORIZATION - OK Failed");
        			e.printStackTrace();
            		tcStatus = "FAIL";
            		StringWriter strWriter = new StringWriter();
            		e.printStackTrace(new PrintWriter(strWriter));
            		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet,taskCode, row, "F",strWriter.toString());
            		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                	{
                		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                	}
        		}
        	}
        	if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
        	{
        		getScreenShot(objExecuteTest.evidencePath, fileName_E);
        		
        	}
        	
        	if(objectName.equalsIgnoreCase("cashDeposit.oK"))
			{
				//Thread.sleep(500);
				Thread.sleep(500);
				loggerUI.info("Object Value: " + objProperties.getProperty(objectName));
			}
        	try
        	{
				wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
				try
				{
					element = driver.findElement(UIOperations.getObject(objProperties,objectName,objectType));
				}
				catch(Exception e)
				{
					loggerUI.info("Unable to Find Click Button ");
					loggerUI.info("Exception In Finding CLICK BUTTON");
					e.printStackTrace();
				}
				try
				{
					act.moveToElement(element).click().build().perform();
					//Thread.sleep(1000);
					act.sendKeys(Keys.TAB).build().perform();
				}
				catch(Exception e)
				{
					loggerUI.info("Unable to Click - Click Button ");
					e.printStackTrace();
				}
        		
				while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		if(loginAlertCheck)		//	If "Invalid User" Alert is thrown, loginCheckFlag is changed to FALSE
                		{
                			ExecuteTest.loginCheckFlag = false;
                			loggerUI.info("Invalid User - 2");
                			break;
                		}
                	}
                	
                }
				if(loginAlertCheck)			//	If "Invalid User" Alert is thrown, update the report as "F"
				{
					objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage);	//	If "Invalid User" Alert is thrown, update the respective test steps to "F"
					loggerUI.info("Excel updated - loginAlertCheck - F and TC_Status Array - False");
					loggerUI.info(objectName + " --> Button Click - PASS ---> Updated in the report");
	        		t_Status.add("False");
	        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                	{
                		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                	}
	        		
				}
				else if(!alertMessage.equalsIgnoreCase(""))		//	If the alert msg contains the word "Invalid" or "Not", update respective test steps to "F"
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Set Text - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info("Click Button - Fail");
            			loggerUI.info(objectName + " --> Click Button - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Click Button - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P","");
            			System.out.println("Set Text - Pass");
                		loggerUI.info(objectName + " --> Click Button - PASS ---> Updated in the report");
                		t_Status.add("True");
                		
        			}
        			
        		}
				else
				{
					objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P","");	//  Alert_Msg:  Emptied the alert message value
					loggerUI.info(objectName + " --> Button Click - PASS ---> Updated in the report");
					System.out.println("Excel updated - loginAlertCheck - P and TC_Status Array - True");
	        		t_Status.add("True");
	        	}
/****************** Added to capture Success Message - Starts ************************/					
				/*{
					waitForFrame("bottom",objectType);
					if(driver.findElements(By.xpath("//input[@value=' Server Interaction Successful.']")).size() != 0)
					{
						objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, ExecuteTest.sheetName_R, row, "P",alertMessage);	
						loggerUI.info(objectName + " --> Button Click - PASS ---> Updated in the report");
						System.out.println("Excel updated - loginAlertCheck - P and TC_Status Array - True");
		        		t_Status.add("True");
					}
					else
					{
						objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, ExecuteTest.sheetName_R, row, "P","Server Interaction Successful.");	
						loggerUI.info(objectName + " --> Button Click - PASS ---> Updated in the report");
						System.out.println("Excel updated - loginAlertCheck - P and TC_Status Array - True");
		        		t_Status.add("True");
					}
				}*/
/****************** Added to capture Success Message - Ends ************************/				
        	}
        	catch (Exception ex)
        	{
        		loggerUI.info("CLICK BUTTON Failed");
        		loggerUI.error(objectName + " --> Click Button Failed");
        		ex.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.error(objectName + " --> Button Click - FAIL ---> Updated in the report");
        		StringWriter strWriter = new StringWriter();
        		ex.printStackTrace(new PrintWriter(strWriter));
        		if(alertMessage.contains("Invalid user"))
        		{
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage);
        			
        		}
        		else
        		{
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		}
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
           	}
            break;
            
        case "CLICKLINK":
            //	Perform click on Link
           	loggerUI.info("CLICK LINK_SLEEP: " + i);
        	loggerUI.info("Window Title before Environment click: " + driver.getTitle());
        	if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
        	{
        		getScreenShot(objExecuteTest.evidencePath, fileName_E);
        	}
        	while(isAlert()==true)
            {
            	if(alertFlag)
            	{
            		alertMessage = getCancelAlert();
            		loggerUI.info("Alert - Cancel Clicked ");
            	}
            	else
            	{
            		alertMessage = getAcceptAlert();
            		loggerUI.info("Alert - OK Clicked ");
            	}
            }
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
        		element = driver.findElement(UIOperations.getObject(objProperties,objectName,objectType));
        		act.moveToElement(element).click().sendKeys(Keys.TAB).build().perform();
        		Thread.sleep(500);
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
        		if(!alertMessage.equalsIgnoreCase(""))
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Click Link - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info("Click Link - Fail");
            			loggerUI.info(objectName + " --> Click Link - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Click Link - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            			System.out.println("Click Link - Pass");
                		loggerUI.info(objectName + " --> Click Link - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> Click Link - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
        			System.out.println("Click Link - Pass");
            		loggerUI.info(objectName + " --> Click Link - PASS ---> Updated in the report");
            		t_Status.add("True");
            	}

        	}
        	
            catch(Exception e)
        	{
            	loggerUI.error(objectName + " --> Click Link Failed");
            	e.printStackTrace();
            	if(objectName.equalsIgnoreCase("indexPage.login"))
            	{
            		t_Status.add("True");
            		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                	loggerUI.info(objectName + " --> Click Link_Exception - PASS ---> Updated in the report");
            	}
            	else
            	{
            		t_Status.add("False");
            		loggerUI.info(objectName + " --> Click Link - FAIL ---> Updated in the report");
                	StringWriter strWriter = new StringWriter();
            		e.printStackTrace(new PrintWriter(strWriter));
            		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
            		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                	{
                		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                	}
            	}
        	}
        	
            break;
            
        case "SETTEXT":
        	System.out.println("Value: " + value);
        	while(isAlert()==true)
            {
            	if(alertFlag)
            	{
            		alertMessage = getCancelAlert();
            		loggerUI.info("Alert - Cancel Clicked ");
            	}
            	else
            	{
            		alertMessage = getAcceptAlert();
            		loggerUI.info("Alert - OK Clicked ");
            	}
            }
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
				element = driver.findElement(UIOperations.getObject(objProperties,objectName,objectType));
        		t_Status.add("True");
        		loggerUI.info(objectName + " -->  for Set Text Identified");
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
        	}
        	catch(Exception e)
        	{
    			e.printStackTrace();
    			loggerUI.error(objectName + " --> Set Text - FAIL ---> Updated in the report");
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage);
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
			try
        	{
        		element.clear();
        		element.sendKeys(value,Keys.TAB);	//	Added for FCRC - 8054
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
        		if(!alertMessage.equalsIgnoreCase(""))
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Set Text - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info(objectName + " --> Set Text - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Set Text - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> Set Text - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> Set Text - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> Set Text - PASS ---> Updated in the report");
            		t_Status.add("True");
            	}
        	}
			catch (Exception e)
        	{
				loggerUI.error(objectName + " --> Set Text - PASS");
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.error(objectName + " --> Set Text - FAIL ---> Updated in the report");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
            break;
            
        case "DROPDOWNSELECTVALUE":
        	while(isAlert()==true)
            {
            	if(alertFlag)
            	{
            		alertMessage = getCancelAlert();
            		loggerUI.info("Alert - Cancel Clicked ");
            	}
            	else
            	{
            		alertMessage = getAcceptAlert();
            		loggerUI.info("Alert - OK Clicked ");
            	}
            	
            }
            //	Selecting value in the drop down
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
        		element = driver.findElement(UIOperations.getObject(objProperties, objectName, objectType));
        		Select ddl = new Select(element);
        		ddl.selectByIndex(0);
        		
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
        		if(!alertMessage.equalsIgnoreCase(""))
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> DropDown Select Value - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info(objectName + " --> DropDown Select Value - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> DropDown Select Value - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> DropDown Select Value - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> DropDown Select Value - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> DropDown Select Value - PASS ---> Updated in the report");
            		t_Status.add("True");
        		}
        	
        	}
            catch(Exception e)
        	{
            	loggerUI.error(objectName + " --> DropDown Select Value - FAIL");
    			e.printStackTrace();
            	tcStatus = "FAIL";
            	t_Status.add("False");
            	loggerUI.error(objectName + " --> DropDown Select Value - FAIL ---> Updated in the report");
            	StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
            break;
        
        case "DROPDOWNSENDVALUE":
            //	Sending value to the drop down  
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
        		driver.findElement(UIOperations.getObject(objProperties, objectName, objectType)).sendKeys(value);
        		
            	while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
            	if(!alertMessage.equalsIgnoreCase(""))
        		{
            		if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> DropDown Send Value - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); 
            			loggerUI.info(objectName + " --> Set Text - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> DropDown Send Value - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> DropDown Send Value - PASS ---> Updated in the report");
                		t_Status.add("True");
        			}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> DropDown Send Value - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> DropDown Send Value - PASS ---> Updated in the report");
            		t_Status.add("True");
        		}
        	}
        	catch(Exception e)
        	{
        		loggerUI.error(objectName + " --> DropDown Send Value - FAIL");
    			e.printStackTrace();
        		loggerUI.error(objectName + " --> DropDown Send Value - FAIL ---> Updated in the report");

        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	break;
        	
        case "OPENAPPLICATION":
            //	Launch Application and goto Login Page
        	try
        	{
            	
                driver.get(objProperties.getProperty(objectName));
                driver.manage().window().maximize();
                driver.manage().timeouts().implicitlyWait(3000, TimeUnit.MILLISECONDS);
                Thread.sleep(i);
               loggerUI.info(objectName + " --> Open Application - PASS");
                objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                loggerUI.info(objectName + " --> Open Application - PASS ---> Updated in the report");
                try
                {
                	t_Status.add("True");
                }
                catch(Exception e)
                {
                	System.out.println("Array Exception");
                	e.printStackTrace();
                }
                
        	}
        	catch(Exception e)
        	{
        		loggerUI.error(objectName + " --> Open Application - FAIL");
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.error(objectName + " --> Open Application - FAIL ---> Updated in the report");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	break;
         
        case "GETTEXT":
            //	Get text of an element
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
        		driver.findElement(UIOperations.getObject(objProperties,objectName,objectType)).getText();
        		loggerUI.info(objectName + " --> Get Text - PASS");
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
        		loggerUI.info(objectName + " --> Get Text - PASS ---> Updated in the report");
        		t_Status.add("True");
        	}
            catch(Exception e)
        	{
            	loggerUI.error(objectName + " --> Get Text - FAIL");
    			e.printStackTrace();
            	tcStatus = "FAIL";
            	t_Status.add("False");
            	loggerUI.error(objectName + " --> Get Text - FAIL ---> Updated in the report");
            	StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
            break;
            
        case "SWITCHFRAME":
            //	Switching to New Frame
        	while(isAlert()==true)
            {
            	if(alertFlag)
            	{
            		alertMessage = getCancelAlert();
            		loggerUI.info("Alert - Cancel Clicked ");
            	}
            	else
            	{
            		alertMessage = getAcceptAlert();
            		loggerUI.info("Alert - OK Clicked ");
            	}
            }
        	try
        	{
        	waitForFrame(objProperties.getProperty(objectName),objectType);
        	}
        	catch (Exception e)
        	{
        		loggerUI.error(objectName + " --> Switch Frame - FAIL ---> Updated in the report");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage);
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	break;

        case "SWITCHWINDOW":
            //	Switching to Available Window

            Thread.sleep(i);
            try
            {
            	switchAvailableWindow();

            	objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            	loggerUI.info(objectName + " --> Switch Window - PASS ---> Updated in the report");
        		t_Status.add("True");
            }
        	catch(Exception e)
            {
        		loggerUI.error(objectName + " --> Switch Window - FAIL");
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.error(objectName + " --> Switch Window - FAIL ---> Updated in the report");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
            }
        	
        	break;
        
        case "SWITCH_REPORT_QUERIES":
        	try{
        	System.out.println("Before clickReportsAndQueries" );
        	clickReportsAndQueries();
        	System.out.println("After clickReportsAndQueries" );
        	System.out.println("Done ====");
        	}
        	catch(Exception e )
        	{
        		e.printStackTrace();
        	}
        break;
        
        
        case "SWITCHNEWWINDOW":

        	String currentWindow = driver.getWindowHandle();

            Thread.sleep(i);
            try
            {
            	switchToNewWindow(currentWindow);
            	objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            	loggerUI.info(objectName + " --> Switch New Window - PASS ---> Updated in the report");
        		t_Status.add("True");
            }
            catch(Exception e)
            {
            	loggerUI.error(objectName + " --> Switch New Window - PASS");
    			e.printStackTrace();
            	tcStatus = "FAIL";
            	t_Status.add("False");
            	loggerUI.error(objectName + " --> Switch New Window - FAIL ---> Updated in the report");
            	StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
            }
                	
        case "SWITCHWINDOWBYTITLE":
            //	Switching to New Window By Title
        	try
        	{
        		System.out.println("Window 1 Title: " + driver.getTitle());
        	}
        	catch(Exception e)
        	{
        		System.out.println("Exception in driver");
        		e.printStackTrace();
        	}
        	loggerUI.info("Delay: I " + i);
            try
            {
            	windowSwitchByTitle(objProperties,objectName);
            	objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            	loggerUI.info(objectName + " --> Switch Window By Title - PASS ---> Updated in the report");
        		t_Status.add("True");
            }
            catch(Exception e)
            {
            	loggerUI.error(objectName + " --> Switch Window By Title - FAIL");
    			e.printStackTrace();
            	tcStatus = "FAIL";
            	t_Status.add("False");
            	loggerUI.error(objectName + " --> Switch Window By Title - FAIL ---> Updated in the report");
            	StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
            }
        	
        	break;
        	
        case "SWITCH_BROWSERWINDOW":
        	try{
        	System.out.println("Before SWITCH_BROWSERWINDOW" );
        	windowOpenForPrint();
        	//switchwindowOpenForPrint();
        	System.out.println("After SWITCH_BROWSERWINDOW" );
        	
        	//closePopWindow();
        	System.out.println("Done ====");
        	}
        	catch(Exception e )
        	{
        		e.printStackTrace();
        	}
        break;
        	
        case "SWITCHREPORTSQUERYTAB":
            //	Switching to Reports and query tab
            Thread.sleep(i);
            
        	break;
        	
        
        case "SWITCHWINDOWTAB":
            //	Switching to New Window By Title
            Thread.sleep(i);
            System.out.println("SWITCHWINDOWTAB Object Properties: " + objProperties);
            System.out.println("SWITCHWINDOWTAB Object Name: " + objectName);
            windowSwitchAndPressTab(objProperties,objectName);
        	String parentHandle = driver.getWindowHandle();
//        	switchTest(parentHandle);
//        	switchTest(objProperties,objectName);
//        	windowSwitchByTitle2(objProperties,objectName);
//            switchToTitle(objectName);
        	break;
        	
        case "SWITCHREPORTSANDQUERYTAB":
            //	Switching to New Window By Title
            Thread.sleep(i);
            System.out.println("Reports/Query Object Properties: " + objProperties);
            System.out.println("Reports/Query Object Name: " + objectName);
            switchReportsAndQueryTab(objProperties,objectName);
        	break;
       	
        case "PRESENCEOFELEMENT":
        	//	 Validate Presence of Element in a page
        	if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
        	{
        		getScreenShot(objExecuteTest.evidencePath, fileName_E);
        	}
        	
        	while(isAlert()==true)
            {
            	if(alertFlag)
            	{
            		alertMessage = getCancelAlert();
            	}
            	else
            	{
            		alertMessage = getAcceptAlert();
            		if(loginAlertCheck)		//	If "Invalid User" Alert is thrown, loginCheckFlag is changed to FALSE
            		{
            			ExecuteTest.loginCheckFlag = false;
            			loggerUI.error("Invalid User");
            			break;
            		}
            	}
            	if(!alertMessage.equalsIgnoreCase(""))
        		{
            		if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Set Number - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info(objectName + " --> Set Number - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Set Number - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> Set Number - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
            	
            }
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  
        		if(driver.findElement(UIOperations.getObject(objProperties, objectName, objectType))!=null)
            	{
            		loggerUI.info("Element Present");
            		if(ExecuteTest.loginCheckFlag == false)
            		{
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage);
                		t_Status.add("False");
                		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
            		}
            		else
            		{
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		t_Status.add("True");
            		}
            	}
            	else
            	{
            		loggerUI.info("Element Not Present");
            	}
        	}
        	catch (Exception e)
        	{
        		loggerUI.error("PRESENCE OF ELEMENT Failed");
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		e.printStackTrace();
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage);
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	
        	break;
        
        case "CORPCIFLOGOUT":
        	//	 Logout for CorpCIF from Welcome page	
        	try
        	{
        		WebElement rootMenu = driver.findElement(By.id("rootMenu"));
            	Actions action = new Actions(driver);
    			action.moveToElement(rootMenu).click().perform();
        		WebElement logoutLink = driver.findElement(By.xpath("//a[@id='logout' and @onclick='logoutUser()']"));
        		logoutLink.click();
        		Thread.sleep(100);
        		Thread.sleep(100);
        		driver.findElement(By.linkText("Close")).click();
        		Thread.sleep(100);
        		driver.quit();
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
        		t_Status.add("True");
        	}
        	catch (Exception e)
        	{
        		loggerUI.error("CORPCIF LOGOUT Failed");
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}

        	break;
        
        case "REOPENAPPLICATION":

        	System.out.println("URL2: " + value);
        	try
        	{
        		driver.quit();
        		DesiredCapabilities cap = new DesiredCapabilities();
                cap.setJavascriptEnabled(true);
                System.setProperty("webdriver.ie.driver", "D:\\VF_Polaris\\966031\\20140420\\WAT_Tool\\Configuation\\Utils\\IEDriverServer.exe");  //  Change 8: Update the path of IEDriverServer
            	driver = new InternetExplorerDriver();
            	driver.get(objProperties.getProperty(value));
            	driver.manage().window().maximize();
            	driver.manage().timeouts().implicitlyWait(3000, TimeUnit.MILLISECONDS);
            	objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
        	}
        	catch(Exception e)
        	{
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	break;
            
        case "REFRESHLOGIN":

        	Thread.sleep(i);	
        	try
        	{
        		driver.navigate().refresh();
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
        		t_Status.add("True");
        	}
        	catch(Exception e)
        	{
        		e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	break;
        
        case "SELECTRADIOBUTTON":
        	while(isAlert()==true)
            {
            	if(alertFlag)
            	{
            		alertMessage = getCancelAlert();
            		loggerUI.info("Alert - Cancel Clicked ");
            	}
            	else
            	{
            		alertMessage = getAcceptAlert();
            		loggerUI.info("Alert - OK Clicked ");
            	}
            }
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
        		driver.findElement(UIOperations.getObject(objProperties,objectName,objectType)).click();
        		
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
        		if(!alertMessage.equalsIgnoreCase(""))
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Select Radio Button - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info(objectName + " --> Set Text - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Select Radio Button - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> Select Radio Button - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> Set Text - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> Set Text - PASS ---> Updated in the report");
            		t_Status.add("True");
        		}
        	}
        	catch(Exception e)
        	{
        		loggerUI.error("SELECT RADIO BUTTON Failed ");
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.info("SELECT RADIO BUTTON Result Updated as Fail");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	break;
        
        case "SETNUMBER":

        	while(isAlert()==true)
     		{
     			if(alertFlag)
     			{
     				alertMessage = getCancelAlert();
     				loggerUI.info("Alert - Cancel Clicked ");
     			}
     			else
     			{
     				alertMessage = getAcceptAlert();
     				loggerUI.info("Alert - OK Clicked ");
     			}
     		}
        	try
        	{
/************************************* Below code works fine for setting number value ********************************/        	
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType)));  //  Added_WebDriverWait
            	WebElement elementNumber = driver.findElement(UIOperations.getObject(objProperties,objectName,objectType));
            	js = ((JavascriptExecutor) driver);

            	if(objectName.equals("welcomePage.taskCode"))
            	{
            		js.executeScript("arguments[0].value=" + value + ";", elementNumber);
            	}
            	else
            	{
            		loggerUI.info("Value: " + value);
            		elementNumber.sendKeys(value,Keys.TAB);
            	}
            	if(objectName.equalsIgnoreCase("misCustCredit.glAccountNo"))
            	{
            		//Thread.sleep(1000);
            		Thread.sleep(100);
            	}
            	objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            	loggerUI.info("SET NUMBER Passed ");
/************************************* Above code works fine for setting number value ********************************/ 
        		t_Status.add("True");
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                	}
                }
        		if(!alertMessage.equalsIgnoreCase(""))
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Set Number - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
            			loggerUI.info(objectName + " --> Set Number - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Set Number - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> Set Number - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> Set Number - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> Set Number - PASS ---> Updated in the report");
            		t_Status.add("True");
        		}
        	}
        	catch(Exception e)
        	{
        		loggerUI.info("Unable to Find element for SET NUMBER ---> " + "ObjectName: " + objectName + " ObjectType: " + objectType);
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.error("SET NUMBER Failed ");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
     	
        	break;
        	
        case "SETDATE":
        	if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
        	{
        		getScreenShot(objExecuteTest.evidencePath, fileName_E);
        	}
        	try
        	{
        		wait.until(ExpectedConditions.visibilityOfElementLocated(UIOperations.getObject(objProperties,objectName,objectType))); 
        		element = driver.findElement(UIOperations.getObject(objProperties,objectName,objectType));
            	/*String sDate1="2016/12/25";
            	SimpleDateFormat formatter1=new SimpleDateFormat("yyyy/MM/dd");
            	Date date1=formatter1.parse(sDate1);
            	System.out.println("sDate1: " + sDate1);*/
            	js = ((JavascriptExecutor) driver);
            	js.executeScript("document.getElementById('founddate').setAttribute('value','2016/12/25')");
            	
        		while(isAlert()==true)
                {
                	if(alertFlag)
                	{
                		alertMessage = getCancelAlert();
                		loggerUI.info("Alert - Cancel Clicked ");
                	}
                	else
                	{
                		alertMessage = getAcceptAlert();
                		loggerUI.info("Alert - OK Clicked ");
                	}
                }
        		if(!alertMessage.equalsIgnoreCase(""))
        		{
        			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
        			{
        				loggerUI.info(objectName + " --> Set Date - FAIL");
        				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); 
            			loggerUI.info(objectName + " --> Set Text - FAIL ---> Updated in the report");
            			t_Status.add("False");
            			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                    	{
                    		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                    	}
        			}
        			else
        			{
        				loggerUI.info(objectName + " --> Set Date - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                		loggerUI.info(objectName + " --> Set Text - PASS ---> Updated in the report");
                		t_Status.add("True");
                	}
        			
        		}
        		else
        		{
        			loggerUI.info(objectName + " --> Set Date - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> Set Date - PASS ---> Updated in the report");
            		t_Status.add("True");
        		}
        	}
        	catch(Exception e)
        	{
    			e.printStackTrace();
        		tcStatus = "FAIL";
        		t_Status.add("False");
        		loggerUI.error("SET DATE Failed ");
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	
        	if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
        	{
        		getScreenShot(objExecuteTest.evidencePath, fileName_E);
        	}
        	break;
        
		case "AUTHORIZATION":
			System.out.println("FCRC Authorization: Object Name: " + objectName + "Object value: " + objProperties.getProperty(objectName));
			while(isAlert()==true)
			{
				if(alertFlag)
				{
					alertMessage = getCancelAlert();
					loggerUI.info("Alert - Cancel Clicked ");
				}
				else
				{
					alertMessage = getAcceptAlert();
					loggerUI.info("Alert - OK Clicked ");
				}
			}
/************************************/	
			if(!alertMessage.equalsIgnoreCase(""))
    		{
    			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
    			{
    				loggerUI.info(objectName + " --> Authorization - FAIL");
    				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
        			loggerUI.info(objectName + " --> Authorization - FAIL ---> Updated in the report");
        			t_Status.add("False");
        			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                	{
                		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                	}
    			}
    			else
    			{
    				loggerUI.info(objectName + " --> Authorization - PASS");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            		loggerUI.info(objectName + " --> Authorization - PASS ---> Updated in the report");
            		t_Status.add("True");
            	}
    			
    		}
    		else
    		{
    			loggerUI.info(objectName + " --> Authorization - PASS");
    			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
    			loggerUI.info(objectName + " --> Authorization - PASS ---> Updated in the report");
        		t_Status.add("True");
    		}
/************************************/			
        	try
        	{
        		if(driver.findElements(By.linkText("Local Auth")).size()!=0)
        		{
        			element = driver.findElement(UIOperations.getObject(objProperties,objectName,objectType));
            		act.moveToElement(element).click().sendKeys(Keys.TAB).build().perform();
            		String author = value.split("_")[0];
            		String pw = value.split("_")[1];
            		driver.findElement(By.id("UserID")).sendKeys(author);
            		driver.findElement(By.id("Passwd")).sendKeys(pw);
            		if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
                	{
                		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                	}
            		WebElement clickButton = driver.findElement(By.id("btnOk"));
            		act.moveToElement(clickButton).click().sendKeys(Keys.TAB).build().perform();
            		while(isAlert()==true)
                    {
                    	if(alertFlag)
                    	{
                    		alertMessage = getCancelAlert();
                    	}
                    	else
                    	{
                    		alertMessage = getAcceptAlert();
                    	}
                    	
                    }
            		
            		if(!alertMessage.equalsIgnoreCase(""))
            		{
            			if(alertMessage.contains("Invalid") || alertMessage.contains("Not") || alertMessage.contains("invalid") || alertMessage.contains("not") || alertMessage.contains("Rejected") || alertMessage.contains("rejected"))
            			{
            				loggerUI.info(objectName + " --> Authorization - FAIL");
            				objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",alertMessage); //  Alert_Msg:  Changed to "P" from "F" and emptied the alert message value
                			loggerUI.info(objectName + " --> Authorization - FAIL ---> Updated in the report");
                			t_Status.add("False");
                			if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
                        	{
                        		getScreenShot(objExecuteTest.evidencePath, fileName_E);
                        	}
            			}
            			else
            			{
            				loggerUI.info(objectName + " --> Authorization - PASS");
                			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
                			loggerUI.info(objectName + " --> Authorization - PASS ---> Updated in the report");
                    		t_Status.add("True");
                    	}
            			
            		}
            		else
            		{
            			loggerUI.info(objectName + " --> Authorization - PASS");
            			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
            			loggerUI.info(objectName + " --> Authorization - PASS ---> Updated in the report");
                		t_Status.add("True");
            		}
        		}
        		else
        		{
        			loggerUI.info("Local Auth Not Found");
        			loggerUI.info("Authorization Not Found - Success");
        			objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
        			loggerUI.info("AUTHORIZATION Result Upated as Pass");
        		}
        	}
        	catch(Exception e)
        	{
        		loggerUI.error("AUTHORIZATION Failed");
        		loggerUI.info("AUTHORIZATION Result Upated as Failed");
        		t_Status.add("False");
        		e.printStackTrace();
        		StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
        	}
        	
        	break;
        	
		case "ALERTCANCEL":
//			To click "Cancel" in the Alert
			while(isAlert()==true)
            {
				getCancelAlert();
            }
		
		case "CLOSEBROWSER":
			try
			{
				loggerUI.info("Driver 1: " + driver);
				driver.quit();
	        	Thread.sleep(500);
	        	loggerUI.info("All Browsers - Closed");
	        	System.setProperty("webdriver.ie.driver", webdriveLocation);  //  Change 8: Update the path of IEDriverServer
	    		driver = new InternetExplorerDriver();
	    		loggerUI.info("New Browser Launched");
	    		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "P",alertMessage);
	    		loggerUI.info("Browsers - Closed Result Updated as Pass ");
	    		t_Status.add("True");
	    		loggerUI.info("Driver 2: " + driver);
			}
			catch(Exception e)
			{
				t_Status.add("False");
				loggerUI.error("Browsers - Closed Result Updated as Fail");
				StringWriter strWriter = new StringWriter();
        		e.printStackTrace(new PrintWriter(strWriter));
        		objUpdateResult.writeOutputWorkBook(xssfWorkbook, actSheet, taskCode, row, "F",strWriter.toString());
        		if(Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE")))
            	{
            		getScreenShot(objExecuteTest.evidencePath, fileName_E);
            	}
			}
			break;	
			
        default:
            break;
        }
    }
    
    

	/**
     * Find element BY using object type and value
     * @param p
     * @param objectName
     * @param objectType
     * @return
     * @throws Exception
     */
    private static By getObject(Properties p,String objectName,String objectType) throws Exception
    {
    	// Find by ID
    	if(objectType.equalsIgnoreCase("ID"))
        {
            if(objectName.equals("welcomePage.loginCheck"))
            	Thread.sleep(500);
            return By.id(p.getProperty(objectName));
        }
    	//	Find by name
        else if(objectType.equalsIgnoreCase("NAME"))
        {
            loggerUI.info("Object Name: " + objectName);
            
            try
            {
            	return By.name(p.getProperty(objectName));
            }
            catch(Exception e)
            {
            	e.printStackTrace();
            	switchAvailableWindow();
            	return By.name(p.getProperty(objectName));
            }
        }
        //	Find by class
        else if(objectType.equalsIgnoreCase("CLASSNAME"))
        {
            return By.className(p.getProperty(objectName));
            
        }
        //	Find by link
        else if(objectType.equalsIgnoreCase("LINKTEXT"))
        {
            return By.linkText(p.getProperty(objectName));
        }
        //	Find by partial link text
        else if(objectType.equalsIgnoreCase("PARTIALLINKTEXT"))
        {
            return By.partialLinkText(p.getProperty(objectName));
        }
    	//	Find by tag name
        else if(objectType.equalsIgnoreCase("TAGNAME"))
        {
            return By.tagName(p.getProperty(objectName));
        }
    	//	Find by Xpath
        else if(objectType.equalsIgnoreCase("XPATH"))
        {
            if(p.getProperty("//div[@class='cssmenu']/following::li[contains(text(),'Welcome!')]") != null)
            {
            	Thread.sleep(500);
            }
            return By.xpath(p.getProperty(objectName));
        }
        //	Find by css
        else if(objectType.equalsIgnoreCase("CSS"))
        {
          return By.cssSelector(p.getProperty(objectName));
       }
       else
        {
    	    loggerUI.info("Wrong Object Type: " + objectType);
    	    loggerUI.error("Invalid Object Type");
            throw new Exception("Wrong object type");
        }
    }
    
    public static String getDateObject(Properties p,String objectName,String objectType)
    {
    	return p.getProperty(objectName);
   }
    // Switching Window
    public static void switchToNewWindow(String parentWindow) throws Exception 
    {
		try 
		{
			Set<String> handles = driver.getWindowHandles();
			loggerUI.info("Windows Count: " + handles.size());
			for (String windowHandle : handles) {
				loggerUI.info("Window Handles: " + driver.getWindowHandle());
				if ((!windowHandle.equals(parentWindow))) 
				{
					try {
						driver.switchTo().window(windowHandle);

						if (driver.getTitle().contains("Certificate Error")) 
						{
							loggerUI.info("Certificate error recieved, by passing certificate error");
							Thread.sleep(1000);
							driver.navigate().to("javascript:document.getElementById('overridelink').click()");
							break;
						}
					} 
					catch (Exception e) 
					{
						e.printStackTrace();
					}
				}
			}
		} catch (NoSuchWindowException e) {
			loggerUI.error("Switch To New Window - Failed");
		}
	}
    // 	Switch to the Available Window
    public static void switchAvailableWindow() throws Exception 
    {
		try 
		{
			Set<String> handles = driver.getWindowHandles();
			System.out.println("Windows Count: " + handles.size());
			for (String windowHandle : handles) 
			{
				driver.switchTo().window(windowHandle);
				System.out.println("Window Title 2: " + driver.getTitle());
				/*try 
				{
					if (driver.getTitle().contains("Certificate Error")) {

						Thread.sleep(1000);
						driver.navigate().to("javascript:document.getElementById('overridelink').click()");
						break;
					}
				} 
				catch (Exception e) 
				{
					e.printStackTrace();
				}*/
			}
		} 
		catch (NoSuchWindowException e) 
		{
			loggerUI.info("Switching to window failed");
			loggerUI.error("Switch Available Window - Failed");
		}
	}
    
    //	Window Switch using window title
    public boolean windowSwitchByTitle(Properties p,String objectName) throws Exception
    {
     boolean switched = false;
     
     do 
     {
       Thread.sleep(1000);
       Set<String> handles =  driver.getWindowHandles();
            System.out.println("Windows Count: " + handles.size());
            for(String windowHandle  : handles)
              {
                      driver.switchTo().window(windowHandle);  
                      System.out.println("Switched to "+driver.getTitle());
                      System.out.println("Object Name: " + p.getProperty(objectName));
                     System.out.println();
                     if(driver.getTitle().contains("CIF"))
                     {
                    	 System.out.println("\"CIF: Finally Switched to "+driver.getTitle());
                    	 switched = true;
                    	 break;
                     }
                     System.out.println("CHECK 3");
                     if((driver.getTitle().contains(p.getProperty(objectName))))
                     {         
                    	 System.out.println("CHECK 4");
                    	 System.out.println("\"Finally Switched to "+driver.getTitle());
                    	 switched = true;
                    	 break;
                     }     
                  }
            return switched;
      } while (!switched);
    } 
    
    
    public void clickReportsAndQueries() throws Exception
    {
    	System.out.println("Inside the windowOpenForPrint () method Started : ");
    	try {
       	Robot robot = new Robot();
    	Thread.sleep(2000);
    	robot.setAutoDelay(100);
    	robot.setAutoWaitForIdle(true);
    	System.out.println("windowOpenForPrint: " + robot);
    	robot.keyPress(KeyEvent.VK_TAB);
    	robot.keyRelease(KeyEvent.VK_TAB);
    	System.out.println("1 st Tab pressed and released " );
    	robot.keyPress(KeyEvent.VK_TAB);
    	robot.keyRelease(KeyEvent.VK_TAB);
    	System.out.println("2nd Tab pressed and released " );
    	robot.keyPress(KeyEvent.VK_TAB);
    	robot.keyRelease(KeyEvent.VK_TAB);
    	System.out.println("3rd Tab pressed and released " );
    	robot.keyPress(KeyEvent.VK_TAB);
    	robot.keyRelease(KeyEvent.VK_TAB);
    	System.out.println("4th Tab pressed and released " );
    	}
    	catch(NoSuchWindowException exp)
    	{
    		exp.printStackTrace();
    		System.out.println("Switching to clickReportsAndQueries failed");
    	}
    }
    
    
    public void windowOpenForPrint() throws Exception
    {
    	System.out.println("Inside the windowOpenForPrint () method Started : ");
    	try {
       	Robot robot = new Robot();
    	Thread.sleep(2000);
    	robot.setAutoDelay(100);
    	robot.setAutoWaitForIdle(true);
    	System.out.println("windowOpenForPrint: " + robot);
    	robot.keyPress(KeyEvent.VK_ALT);
    	robot.keyPress(KeyEvent.VK_O);
    	System.out.println("Key Pressed: " );
  //  	Thread.sleep(1000);
    	robot.keyRelease(KeyEvent.VK_ALT);
    	robot.keyRelease(KeyEvent.VK_O);
    	System.out.println("Key Released: " );
 //   	Thread.sleep(10000);
    	}
    	catch(NoSuchWindowException exp)
    	{
    		exp.printStackTrace();
    		System.out.println("Switching to windowOpenForPrint failed");
    	}
    }
    public void closePopWindow() throws Exception
    {
    	System.out.println("Inside the closePopWindow () method Started : ");
    	try {
       	Robot robot = new Robot();
    	Thread.sleep(1000);
    //	robot.setAutoDelay(500);
    	System.out.println("closePopWindow: " + robot);
    	
    	System.out.println("Start closing the Pop Up" );
    	Thread.sleep(1000);
    	robot.keyPress(KeyEvent.VK_ALT);
    	robot.keyPress(KeyEvent.VK_SPACE);
    	robot.keyPress(KeyEvent.VK_C);
    	robot.delay(500);
    	robot.keyRelease(KeyEvent.VK_ALT);
    	robot.keyRelease(KeyEvent.VK_SPACE);
    	robot.keyRelease(KeyEvent.VK_C);
    	System.out.println("Pop Up closed ");
    	Thread.sleep(10000);
    	}
    	catch(NoSuchWindowException exp)
    	{
    		exp.printStackTrace();
    		System.out.println("Switching to closePopWindow failed");
    	}
    }
    
    public boolean switchwindowOpenForPrint() throws Exception {
    	System.out.println("switchwindowOpenForPrint");
		boolean switched = false;
		try {
			Set<String> handles = driver.getWindowHandles();
			for (String windowHandle : handles) 
			{
				System.out.println("switchwindowOpenForPrint Window Handle Name: " + handles.toString());
    			System.out.println("Current Window Handle: " + driver.getWindowHandle());
    			System.out.println("Window Handle Loop: " + windowHandle);
    			System.out.println("Before Switch Title: "+driver.getTitle());
    			Robot robot = new Robot();  
    			//robot.keyPress(KeyEvent.VK_ENTER);
    			
    			robot.keyPress(KeyEvent.VK_ALT);
    	    	robot.keyPress(KeyEvent.VK_O);
    	    	robot.keyRelease(KeyEvent.VK_ALT);
    	        robot.keyRelease(KeyEvent.VK_O);
    	        
		//		driver.switchTo().window(windowHandle);
			//	System.out.println("Switched to " + driver.getTitle());
			/*	if (driver.getTitle().equals(title.trim())) {
					System.out.println("Finally Switched to "
							+ driver.getTitle());
					Thread.sleep(500);
					switched = true;
					break;
				}*/
			}
		} catch (NoSuchWindowException e) {
			System.out.println("Switching to switchwindowOpenForPrint failed");
			return switched;
		}
		return switched;
	}
    public boolean windowSwitchAndPressTab(Properties p,String objectName) throws Exception
    {
     boolean switched = false;
     System.out.println("windowSwitchAndPressTab" );
     do 
     {
       Thread.sleep(1000);
       Set<String> handles =  driver.getWindowHandles();
            System.out.println("Windows Count: " + handles.size());
            for(String windowHandle  : handles)
              {
    			System.out.println("Window Handle Name: " + handles.toString());
    			System.out.println("Current Window Handle: " + driver.getWindowHandle());
    			System.out.println("Window Handle Loop: " + windowHandle);
    			System.out.println("Before Switch Title: "+driver.getTitle());
    			/*********************************************************/
    			Robot robot = new Robot(); 
    			robot.keyPress(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			
    			/*********************************************************/
    			 driver.switchTo().window(windowHandle);  
    			System.out.println("After Switch Title: "+driver.getTitle());
    			System.out.println("Current Window Handle: " + driver.getWindowHandle());
              System.out.println("Object Name: " + p.getProperty(objectName));
              driver.manage().window().maximize();
                    
           }
           return switched;
      } while (!switched);
    }
    
    public boolean switchReportsAndQueryTab(Properties p,String objectName) throws Exception
    {
     boolean switched = false;
     System.out.println("switchReportsAndQueryTab" );
     do 
     {
       Thread.sleep(1000);
       Set<String> handles =  driver.getWindowHandles();
            System.out.println("Windows Count: " + handles.size());
            for(String windowHandle  : handles)
              {
    			System.out.println("Reports/Query :Window Handle Name: " + handles.toString());
    			System.out.println("Reports/Query :Current Window Handle: " + driver.getWindowHandle());
    			System.out.println("Reports/Query :Window Handle Loop: " + windowHandle);
    			System.out.println("Reports/Query :Before Switch Title: "+driver.getTitle());
    			/*********************************************************/
    			Robot robot = new Robot(); 
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			robot.keyPress(KeyEvent.VK_TAB);
    			robot.keyRelease(KeyEvent.VK_TAB);
    			Thread.sleep(300);
    			
    			/*********************************************************/
    			 driver.switchTo().window(windowHandle);  
    		;
              driver.manage().window().maximize();
                    
           }
           return switched;
      } while (!switched);
    }

    public boolean windowSwitchByTitle2(Properties p,String objectName) throws Exception
    {
     boolean switched = false;
     System.out.println("Object Name: " + p.getProperty(objectName));
     switchAvailableWindow();
     try
     {
    	 System.out.println("Before switching - Window Title: " + driver.getTitle());
     }
     catch(Exception e)
     {
    	 e.printStackTrace();
     }
     do {
       Thread.sleep(2000);
       Set<String> handles =  driver.getWindowHandles();
              System.out.println("Handles: " + handles);
            for(String windowHandle  : handles)
              {
             //Thread.sleep(10000);
              /*if(!(windowHandle.equals(parentWindow)))
                {         
               */
                       driver.switchTo().window(windowHandle);  
                     System.out.println("Switched to "+driver.getTitle());
//                     Thread.sleep(2000);
                     
                     if(driver.getTitle().contains("CIF"))
                     {
                    	 System.out.println("\"CIF: Finally Switched to "+driver.getTitle());
                    	 switched = true;
                    	 break;
                     }
                     if((driver.getTitle().contains(p.getProperty(objectName))))
                     {          
//                  	Thread.sleep(1000);
                    	 System.out.println("\"Finally Switched to "+driver.getTitle());
                    	 switched = true;
                    	 break;
                     }     
                  }
//                }
                  
            return switched;
      } while (!switched);
    }
    public boolean switchToTitle(String title) throws Exception {
		boolean switched = false;
		try {
			Set<String> handles = driver.getWindowHandles();
			for (String windowHandle : handles) 
			{
				System.out.println("Window Handle Name: " + handles.toString());
    			System.out.println("Current Window Handle: " + driver.getWindowHandle());
    			System.out.println("Window Handle Loop: " + windowHandle);
    			System.out.println("Before Switch Title: "+driver.getTitle());
    			Robot robot = new Robot();  
    			robot.keyPress(KeyEvent.VK_ENTER); 
				driver.switchTo().window(windowHandle);
				System.out.println("Switched to " + driver.getTitle());
				if (driver.getTitle().equals(title.trim())) {
					System.out.println("Finally Switched to "
							+ driver.getTitle());
					Thread.sleep(500);
					switched = true;
					break;
				}
			}
		} catch (NoSuchWindowException e) {
			System.out.println("Switching to window failed");
			return switched;
		}
		return switched;
	}
    
    //	Switching Frame
    
    private boolean  waitForFrame(String objectName, String objectType) throws Exception 
    {
		String frame = objectName;
    	boolean switched = false;
		boolean multiframes = frame.contains("_");
		if(frame.equalsIgnoreCase("left_fraMenu"))
		{
			Thread.sleep(1000);
		}
		while(isAlert()==true)
		{
			if(alertFlag)
			{
				getCancelAlert();
				loggerUI.info("Alert - Cancel Clicked ");
			}
			else
			{
				getAcceptAlert();
				loggerUI.info("Alert - OK Clicked ");
			}
		}
		if (!multiframes) 
		{
			switched = false;
			try 
			{
				long startTime = System.currentTimeMillis();
				driver.switchTo().defaultContent();
				WebDriverWait wait = new WebDriverWait(driver, 3);
				wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(frame));
				switched = true;
				loggerUI.info(
						"Switched to " + frame + " Successfully  in " + (System.currentTimeMillis() - startTime));
				loggerUI.info("Switched To New Frame: " + frame + " - Successfully");
			} 
			catch (Exception e) 
			{
				loggerUI.info("frame does not exist: " + frame);
				loggerUI.info("Current Window Title: " + driver.getTitle() );
				e.printStackTrace();
				loggerUI.error("Switch To New Frame: " + frame + " - Failed");
				return false;
			}
		} 
		else if (multiframes) 
		{
			try
			{
				long startTime = System.currentTimeMillis();
				int l = frame.split("_").length;
				driver.switchTo().defaultContent();
				for (int f = 0; f < l; f++) 
				{
					WebDriverWait wait = new WebDriverWait(driver, 2);
					wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(frame.split("_")[f]));
				}
				System.out
						.println("Switched to " + frame + " Successfully  in " + (System.currentTimeMillis() - startTime));

				switched = true;
			}
			catch (Exception e) 
			{
				loggerUI.info("Mutli frame does not exist: " + frame);
				loggerUI.info("Current Window Title: " + driver.getTitle() );
				e.printStackTrace();
				loggerUI.error("Switch To New Frame: " + frame + " - Failed");
				return false;
			}
		}
		loggerUI.info("Switched To Multi Frame: " + frame + " - Successfully");
		return switched;
	}
    //	Screen Shot
    public static void getScreenShot(String filePath_E, String fileName_E) throws HeadlessException, AWTException, IOException, InterruptedException
    {
      	String url = filePath_E+"\\"+fileName_E+"_"+imagedir+"_"+imagecounter+".jpg";
            System.out.println("ScreenShot Taken : " + url);
            BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
            File screenShot = new File(filePath_E +"\\" + fileName_E + "\\" + fileName_E + "_" + imagecounter + ".jpg");
            ImageIO.write(image, "jpg", screenShot);
            imagecounter++ ;
            
    }
    //	File name or Folder Name Reset
	public void tcName(String folderName) 
	{
		fileName_E = folderName;
		loggerUI.info("fileName_R1: " + fileName_E);
		imagecounter=1;
		new File(objExecuteTest.evidencePath + "\\" + fileName_E).mkdirs();
	}

//	Select value from DDL
	public static void selectFromDDL(Properties p,String objectName,String objectType, String inputdata)
    {
        if(!(inputdata==""))
        {
           try
            {
                WebElement element = driver.findElement(getObject(p,objectName,objectType));
            	Select ddl = new Select(element);
                ddl.selectByVisibleText(inputdata);
                loggerUI.info("DDL value Selected: " + inputdata + " - Successfully");
            }
            catch (Exception e)
            {
                loggerUI.error("DDL value Selected: " + inputdata + " - Failed");
            }
        }
        else
        {
        	System.out.println("Input Value found as NULL");
        	loggerUI.error("The input value for the DDL is found as NULL");
        }
    }
	
// Select Radio Button by Name or Xpath
	public static void selectRadio(String proptype, String propvalue, String value) throws Exception {
		if (!proptype.equalsIgnoreCase("name"))
			throw new Exception("RADIO BUTTON property should be NAME");
		if (value != null && value.equalsIgnoreCase("EMPTY")) {
			return;
		}
		java.util.List<WebElement> allradioElements = driver.findElements(By.name(propvalue));
		try {
			for (WebElement element : allradioElements) {
				if (value == null || value.isEmpty()) {
					element.click();
					return;
				}
				if (element.getAttribute("value").equalsIgnoreCase(value)) {
					element.click();
					return;
				}
				// for CorpCif
				else if (value.contains("/")) {
					String[] arr = value.split("/");
					if (element.getAttribute("value").contains(arr[0])) {
						String s = value.split("/")[1];
						int index = Integer.parseInt(s);
						element = allradioElements.get(index);
					}
					element.click();
					return;
				}
			}
		throw new Exception("RADIO BUTTON - Value not found");
		} catch (org.openqa.selenium.NoSuchElementException e) {
		} catch (ElementNotVisibleException e1) {
		} catch (Exception e2) {
			e2.printStackTrace();
		}
	}
//		For Date Setting
	public static  void dateSetter(WebElement element,String date) throws Exception
	{
		try
		{
		loggerUI.info("Date Picker : " + element + " : " + date);
		loggerUI.info("Title before clicking date picker : "  + driver.getTitle());

		if(date == null || date.isEmpty())
			return;

		Set<String> oldWindows = driver.getWindowHandles();
		element.click();

		waitForNoOfWindows(oldWindows.size() + 1);
		Thread.sleep(1000);			//	Changed from 3000 to 9000 for CIB Admin
		switchToNewWindow(oldWindows);
		Thread.sleep(1000);                   //	Added for CIB Admin
		try
		{
			driver.navigate().to("javascript:document.getElementById('overridelink').click()");
		}
		catch (Exception exp)
		{
			exp.printStackTrace();
		}
		try
		{
			((JavascriptExecutor)driver).executeScript("javascript:set_datetime("+String.valueOf(formatter.parse(date).getTime())+",true)");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	
		loggerUI.info("DATE SELECTED: " + date);
		}
		catch(org.openqa.selenium.UnhandledAlertException e)
		{
			Alertcaught();
		}
		catch (Exception e)
		{
			e.printStackTrace();
			throw e;
		}
	}
	//		For Alert Handling
	public static  String Alertcaught()
	{
		String Result= "";
		try
		{
		Alert alert = driver.switchTo().alert();
		Result=alert.getText();
		alert.accept();
		}
		catch (org.openqa.selenium.NoAlertPresentException ex)
		{
			ex.printStackTrace();
		}
		catch(org.openqa.selenium.NoSuchWindowException exp)
		{
			exp.printStackTrace();
		}
		return Result;

	}
	
//	Wait for the respective no of windows to be opened:  For Calendar
	public static void waitForNoOfWindows(int noOfWindows) throws Exception
	{
		boolean windowOpen =  false;
		for (int i = 0; i < 50; i++)
		{
			Set<String> handles =  driver.getWindowHandles();
		   if(handles.size() == noOfWindows)
		   {
				windowOpen = true;
				break;
		   }
		   else
		   {
				Thread.sleep(1000);
		   }
		}
		if(windowOpen == false)
			throw new Exception("Required number of window not open");
	}

//		Switch to new window from the list of old windows mainly for Calendar
	
	public static void switchToNewWindow(Set<String> oldWindows) throws Exception
	{
		Thread.sleep(1000);   // Added for CorpCIF on June-14
			try
			{
				 Set<String> handles =  driver.getWindowHandles();
				 System.out.println("Windows Count: " + handles.size());
				 for(String windowHandle  : handles)
			     {
					 System.out.println("CHECK 1");
			    	   if(!oldWindows.contains(windowHandle))
			           {
			    		   System.out.println("CHECK 2");	
			    		   driver.switchTo().window(windowHandle);
			        	   	System.out.println("Swithced to New Window");
			        	   	try
			    			{
			    				if(driver.getTitle().contains("Certificate Error"))
			    				{
			    					System.out.println("Certificate error recieved, by passing certificate error");
			    					Thread.sleep(1000);
			    					driver.navigate().to("javascript:document.getElementById('overridelink').click()");
			    				}
			    			}
			    			catch(Exception e)
			    			{
			    			}
			        	   	break;
			            }
			      }
			}
			catch (NoSuchWindowException e)
			{
				System.out.println("Switching to window failed");
				e.printStackTrace();
			}
	}
	public static boolean isAlert()
	{
		try
		{
			Thread.sleep(500);
			Alert alert = driver.switchTo().alert();
			if((alert.getText()).contains("Print"))
			{
				alertFlag = true;
			}
			return true;
		}
		catch(Exception e)
		{
			return false;
		}
	}
	public String getAcceptAlert() throws HeadlessException, AWTException, IOException, InterruptedException
	{
		
		Alert alert = driver.switchTo().alert();
		String alertMsg = alert.getText();
		loggerUI.info("Alert Message: " + alertMsg);
		if(alert.getText().equalsIgnoreCase("Invalid user")|| alert.getText().equalsIgnoreCase("Invalid login"))
		{
			loginAlertCheck = true;
			loggerUI.info("Invalid User Alert Occured");
			
		}
		/*if((Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE"))) || (Fail_SnapShot.equalsIgnoreCase("TRUE") && (Pass_SnapShot.equalsIgnoreCase("TRUE"))))
    	{
			getScreenShot(objExecuteTest.evidencePath, fileName_E);
    	}*/
		if(alert.getText().contains("Invalid")|| alert.getText().contains("invalid") || alert.getText().contains("Not") || alert.getText().contains("not"))
		{
			if((Fail_SnapShot.equalsIgnoreCase("TRUE") && !(Pass_SnapShot.equalsIgnoreCase("TRUE"))) || (Fail_SnapShot.equalsIgnoreCase("TRUE") && (Pass_SnapShot.equalsIgnoreCase("TRUE"))))
	    	{
				getScreenShot(objExecuteTest.evidencePath, fileName_E);
				System.out.println("CHECK ONE");
	    	}
		}
		else
		{
			if((Pass_SnapShot.equalsIgnoreCase("TRUE")) && !(Fail_SnapShot.equalsIgnoreCase("TRUE")) || (Pass_SnapShot.equalsIgnoreCase("TRUE") && (Fail_SnapShot.equalsIgnoreCase("TRUE"))))
			{
				getScreenShot(objExecuteTest.evidencePath, fileName_E);
				System.out.println("CHECK TWO");
			}
		}
		alert.accept();
		loggerUI.info("Alert - OK:  Clicked");
		Thread.sleep(100);
		alertMessage = alertMsg;
		return alertMsg;
	}
	
	public String getCancelAlert() throws HeadlessException, AWTException, IOException, InterruptedException
	{
		if(Pass_SnapShot.equalsIgnoreCase("TRUE"))
    	{
			getScreenShot(objExecuteTest.evidencePath, fileName_E);
    	}
		Alert alert = driver.switchTo().alert();
		String alertMsg = alert.getText();
		alert.dismiss();
		loggerUI.info("Alert - CANCEL:  Clicked");
		Thread.sleep(100);
		return alertMsg;
	}
}
