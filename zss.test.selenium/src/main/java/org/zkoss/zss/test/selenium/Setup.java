package org.zkoss.zss.test.selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Properties;

import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;

public class Setup {

	static Properties config = new Properties();
	
	static{
		//load default properties
		File[] configFiles = new File[] {
				new File("./zss.test.selenium.properties"),
				new File("./zss.test.selenium.local.properties")};
		InputStream is = null;
		for(File f:configFiles){
			try {
				if(f.exists()){
					is = new FileInputStream(f);
					config.load(is);
					System.out.println("Load config "+f.getAbsolutePath());
				}
			} catch(Exception x){
				System.err.print(x.getMessage());
			} finally{
				if(is!=null){
					try {
						is.close();
					} catch (IOException e) {}
					is = null;
				}
			}
		}
		//http://chromium.googlecode.com/, http://chromedriver.storage.googleapis.com/index.html
		//System.setProperty("webdriver.chrome.driver",config.getProperty("webdriver.chrome.driver"));
	}
																																																																																																																																																																																	
	public static String getConfig(String key){
		return getConfig(key,"");
	}
	public static String getConfig(String key,String defval){
		return config.getProperty(key, defval);
	}
	public static int getConfigAsInt(String key){
		return getConfigAsInt(key,0);
	}
	public static int getConfigAsInt(String key,int defval){
		String val = getConfig(key,"");
		if("".equals(val)){
			return defval;
		}
		return Integer.parseInt(val);
	}
	public static long getConfigAsLong(String key){
		return getConfigAsInt(key,0);
	}
	public static long getConfigAsLong(String key,long defval){
		String val = getConfig(key,"");
		if("".equals(val)){
			return defval;
		}
		return Long.parseLong(val);
	}
	public static boolean getConfigAsBoolean(String key){
		return getConfigAsBoolean(key,false);
	}
	public static boolean getConfigAsBoolean(String key,boolean defval){
		String val = getConfig(key,"");
		if("".equals(val)){
			return defval;
		}
		return Boolean.parseBoolean(val);
	}	
	
	
	static public WebDriver getDriver() throws MalformedURLException{
		return getChromeDriver();
//		return getFirefoxDriver();
//		return getIEDriver();
	}
	
	private static WebDriver getChromeDriver() {
		ChromeDriver driver = new ChromeDriver();
		// remote chrome driver
//		RemoteWebDriver driver = null;
//		try {
//			driver = new RemoteWebDriver(new URL("http://10.1.3.223:4444/wd/hub"), DesiredCapabilities.chrome());
//		} catch (MalformedURLException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
		
		String val = getConfig("zss.browserSize","0,0,1200,800");
		String[] vals = val.split(",");
		
		driver.manage().window().setPosition(new Point(Integer.parseInt(vals[0]),Integer.parseInt(vals[1])));
		driver.manage().window().setSize(new Dimension(Integer.parseInt(vals[2]), Integer.parseInt(vals[3])));
		return driver;
	}
	
	private static WebDriver getFirefoxDriver() {
//		FirefoxDriver driver = new FirefoxDriver();
		RemoteWebDriver driver = null;
		try {
			driver = new RemoteWebDriver(new URL("http://10.1.3.222:4444/wd/hub"), DesiredCapabilities.firefox());
		} catch (MalformedURLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		String val = getConfig("zss.browserSize","0,0,1200,800");
		String[] vals = val.split(",");
		
		driver.manage().window().setPosition(new Point(Integer.parseInt(vals[0]),Integer.parseInt(vals[1])));
		driver.manage().window().setSize(new Dimension(Integer.parseInt(vals[2]), Integer.parseInt(vals[3])));
		return driver;
	}
	
	private static WebDriver getIEDriver() {
		// remote ie driver
		RemoteWebDriver driver = null;
		try {
			driver = new RemoteWebDriver(new URL("http://10.1.3.168:4444/wd/hub"), DesiredCapabilities.internetExplorer());
		} catch (MalformedURLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		String val = getConfig("zss.browserSize","0,0,1200,800");
		String[] vals = val.split(",");
		
		driver.manage().window().setPosition(new Point(Integer.parseInt(vals[0]),Integer.parseInt(vals[1])));
		driver.manage().window().setSize(new Dimension(Integer.parseInt(vals[2]), Integer.parseInt(vals[3])));
		return driver;
	}

	static public File getScreenshotFolder(){
		File folder = new File(getConfig("zss.screenshotFolder","./screenshot"));
		if(!folder.exists()){
			folder.mkdirs();
		}
		return folder;
	}
	
	static public File getScreenshotFile(String name){
		return new File(getScreenshotFolder(),name);
	}
	
	static public File getResultFolder(){
		File folder = new File(getConfig("zss.resultFolder","./fail"));
		if(!folder.exists()){
			folder.mkdirs();
		}
		return folder;
	}
	
	static public File getResultFile(String name){
		return new File(getResultFolder(),name);
	}
	
	static public Comparator getImageComparator(){
		return new DefaultComparator(Integer.MAX_VALUE, Integer.MAX_VALUE, 10);//entire image
	}

	//public static String getTestSiteUrlBase() {
	//	return getConfig("zss.testSiteBaseUrl","http://localhost:8080/zss.test/");
	//}

	public static long getAuDelay() {
		return getConfigAsLong("zss.auDelay",2000);
	}
	
	public static boolean isJsDebug(){
		return getConfigAsBoolean("zss.jsDebug",false);
	}
	
	public static int getTimeoutL0(){
		return  getConfigAsInt("zss.wait.l0");
	}
	
	public static int getTimeoutL1(){
		return  getConfigAsInt("zss.wait.l1");
	}

	public static int getTimeoutL2(){
		return  getConfigAsInt("zss.wait.l2");
	}
	
	public static int getTimeoutL3(){
		return  getConfigAsInt("zss.wait.l3");
	}
	
	public static int getTimeoutL4(){
		return  getConfigAsInt("zss.wait.l4");
	}
}
