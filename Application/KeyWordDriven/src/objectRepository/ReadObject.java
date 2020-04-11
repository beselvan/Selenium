package objectRepository;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import operations.UIOperations;

public class ReadObject 
{
	Properties propertyObj = new Properties();
	static final Logger loggerRO = LogManager.getLogger(UIOperations.class.getName());
	static LoadPropertySingleton objectLoad = LoadPropertySingleton.getInstance();
	public String properties_fileName = LoadPropertySingleton.configResourceBundle.getString("PropertiesFileName");
	public Properties getObjectRepository() throws IOException
	{
		InputStream stream = new FileInputStream(new File(properties_fileName));    	
		propertyObj.load(stream);
		loggerRO.info("Properties File Loaded ");
		
		return propertyObj;
	}
}
