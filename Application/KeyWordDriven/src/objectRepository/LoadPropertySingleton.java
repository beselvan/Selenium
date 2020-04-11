package objectRepository;

import java.util.ResourceBundle;

/**
 * @author Sreenu Ganta
 * Date    06-October-2017
 */

public class LoadPropertySingleton {
	
	static LoadPropertySingleton objectLoad=null;
	public static ResourceBundle configResourceBundle=null;
	
	private LoadPropertySingleton()
	{
		configResourceBundle=ResourceBundle.getBundle("config");
	}
	public static LoadPropertySingleton getInstance()
	{
		if(objectLoad == null)
		{
			synchronized(LoadPropertySingleton.class){
				if(objectLoad == null)
					objectLoad=new LoadPropertySingleton();
			}
		}
		return objectLoad;
	}
}
