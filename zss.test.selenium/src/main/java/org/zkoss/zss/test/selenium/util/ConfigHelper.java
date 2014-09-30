/* ConfigHelper.java

	Purpose:
		
	Description:
		
	History:
		Tue Sep 18 12:49:43 TST 2014, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

This program is distributed under GPL Version 3.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
 */
package org.zkoss.zss.test.selenium.util;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.util.AbstractSequentialList;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.android.AndroidDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.iphone.IPhoneDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.safari.SafariDriver;

import com.opera.core.systems.OperaDriver;

/**
 * ZSS Selenium Test configuration helper.
 * 
 * @author sam
 * @author jumperchen
 * 
 */
public class ConfigHelper {
	
	private final static String CONFIG_NAME = "zss.test.selenium.properties";
	
	private final static String LOCAL_CONFIG_NAME = "zss.test.selenium.local.properties";

	private final static String[] SUPPORTED_BROWSER = { "firefox", "chrome",
			"opera", "ie", "htmlunit", "iexplore", "ff", "iphone", "android",
			"safari" };

	private final static String ALL_BROWSERS = "all";

	private static List<String> _allBrowsers = new LinkedList<String>();

	private String _client;

	private String _server;

	private String _contextPath;

	private String _timeout;

	private String _delay;

	private String _browser;

	private Properties _properties;

	private long _lastModified;

	private boolean _debuggable;

	private String _zsstheme; 

	private boolean _comparable;

	private int _granularity, _leniency;

	private static ConfigHelper instance = new ConfigHelper();
	
	private Map<String, WebDriver> _driverMap;

	private HashMap<String, String> _driverSetting;

	private HashMap<String, String> _browserRemote;

	private HashMap<String, List<String>> _ignoreMap;

	private volatile boolean _inited;

	private ConfigHelper() {
	}

	public static ConfigHelper getInstance() {
		try {
			if (!instance._inited) {
				instance.init();
				instance._inited = true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return instance;
	}

	public boolean isDebuggable() {
		return _debuggable;
	}
	
	public String getTheme() {
		return _zsstheme;
	}

	public String getClient() {
		return _client;
	}

	public String getServer() {
		return _server;
	}

	public String getContextPath() {
		return _contextPath;
	}

	public String getDelay() {
		return _delay;
	}

	public String getTimeout() {
		return _timeout;
	}

	public String getBrowser() {
		return _browser;
	}

	public long lastModified() {
		return _lastModified;
	}

	public boolean isAllIgnoreCase(String fileName) {
		List<String> ignoreList = _ignoreMap.get(ALL_BROWSERS);
		if (ignoreList == null)
			return false;
		return ignoreList.contains(fileName);
	}

	public boolean isIgnoreCase(String key, String fileName) {
		if (isAllIgnoreCase(fileName))
			return true;

		for (Map.Entry<String, List<String>> me : _ignoreMap.entrySet()) {
			String key2 = me.getKey();
			if (key.matches("^" + key2 + "$")) {
				List<String> ignoreList = me.getValue();
				if (ignoreList != null) {
					if (ignoreList.contains(fileName)) {
						System.out.println("runtime-ignore: " + key + "="
								+ key2);
						return true;
					}
				}
			}
		}
		return false;
	}

	@SuppressWarnings("deprecation")
	private WebDriver getWebDriver(String key, String remotePath) {
		try {
			if (remotePath != null) {
				if ("firefoxdriver".equalsIgnoreCase(key)) {
					return new RemoteWebDriver(new URL(remotePath),
							DesiredCapabilities.firefox());
				} else if ("chromedriver".equalsIgnoreCase(key)) {
					return new RemoteWebDriver(new URL(remotePath),
							DesiredCapabilities.chrome());
				} else if ("safaridriver".equalsIgnoreCase(key)) {
					return new RemoteWebDriver(new URL(remotePath),
							DesiredCapabilities.safari());
				} else if ("internetexplorerdriver".equalsIgnoreCase(key)) {
					return new RemoteWebDriver(new URL(remotePath),
							DesiredCapabilities.internetExplorer());
				} else if ("operadriver".equalsIgnoreCase(key)) {
					return new RemoteWebDriver(new URL(remotePath),
							DesiredCapabilities.opera());
				} else if ("androiddriver".equalsIgnoreCase(key)) {
					return new AndroidDriver(remotePath);
				} else if ("iphonedriver".equalsIgnoreCase(key)) {
					return new RemoteWebDriver(new URL(remotePath),
							DesiredCapabilities.iphone());
				}
			} else {
				if ("firefoxdriver".equalsIgnoreCase(key)) {
					return new FirefoxDriver();
				} else if ("chromedriver".equalsIgnoreCase(key)) {
					return new ChromeDriver();
				} else if ("safaridriver".equalsIgnoreCase(key)) {
					return new SafariDriver();
				} else if ("internetexplorerdriver".equalsIgnoreCase(key)) {
					return new InternetExplorerDriver();
				} else if ("operadriver".equalsIgnoreCase(key)) {
					return new OperaDriver();
				} else if ("androiddriver".equalsIgnoreCase(key)) {
					return new AndroidDriver();
				} else if ("iphonedriver".equalsIgnoreCase(key)) {
					return new IPhoneDriver();
				} else if ("htmlunitdriver".equalsIgnoreCase(key)) {
					return new HtmlUnitDriver(true);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		throw new IllegalArgumentException("Unsopported arguments [" + key + "]");
	}

	public void shutdown() {
		for (WebDriver driver: _driverMap.values()) {
			driver.quit();
		}
		_driverMap.clear();
	}
	/**
	 * 
	 * @param key
	 * @return value :
	 */
	private WebDriver getBrowserFromHolder(String key) {
		key = key.toLowerCase();
		if (_driverSetting.get(key) == null)
			throw new NullPointerException("Null Browser Type String");

		final String driverName = _driverSetting.get(key);
		WebDriver driver = _driverMap.get(driverName);
		if (driver == null) {
			driver = getWebDriver(_driverSetting.get(key), _browserRemote.get(key));
			_driverMap.put(key, driver);
		}
		return driver;
	}

	public List<WebDriver> getBrowsersForLazy(String keys, String caseName) {
		final List<String> drivers = new ArrayList<String>(Arrays.asList(keys
				.split(",")));
		if (drivers.contains(ALL_BROWSERS)) {
			drivers.remove(ALL_BROWSERS);
			drivers.addAll(_allBrowsers);
		}
		//for (Iterator<String> it = drivers.iterator(); it.hasNext();) {
		//	if (isIgnoreCase(it.next(), caseName))
		//		it.remove();
		//}
		List<WebDriver> list = new AbstractSequentialList<WebDriver>() {
			@Override
			public ListIterator<WebDriver> listIterator(int index) {
				return new ItemIter(index, drivers);
			}

			@Override
			public int size() {
				return drivers.size();
			}
		};
		return list;
	}

	private class ItemIter implements ListIterator<WebDriver> {
		private int _j;

		private List<String> _drivers;

		private ItemIter(int index, List<String> drivers) {
			_j = index;
			_drivers = drivers;
		}

		@Override
		public boolean hasNext() {
			return _j < _drivers.size();
		}

		@Override
		public WebDriver next() {
			if (!hasNext())
				throw new NoSuchElementException();
			return getBrowserFromHolder(_drivers.get(_j++));
		}

		@Override
		public void remove() {
			throw new UnsupportedOperationException("Readonly!");
		}

		@Override
		public boolean hasPrevious() {
			return _j > 0;
		}

		@Override
		public WebDriver previous() {
			if (!hasPrevious())
				throw new NoSuchElementException();
			return getBrowserFromHolder(_drivers.get(--_j));
		}

		@Override
		public int nextIndex() {
			return _j;
		}

		@Override
		public int previousIndex() {
			return _j - 1;
		}

		@Override
		public void set(WebDriver e) {
			throw new UnsupportedOperationException("Readonly!");
		}

		@Override
		public void add(WebDriver e) {
			throw new UnsupportedOperationException("Readonly!");
		}
	}

	private void init() throws IOException, Exception {
		if (_driverSetting == null) {
			_driverSetting = new HashMap<String, String>();
		}
		if (_browserRemote == null) {
			_browserRemote = new HashMap<String, String>();
		}
		if (_ignoreMap == null) {
			_ignoreMap = new HashMap<String, List<String>>();
		}
		initProperty();
		//initIgnoreList();
	}

	@SuppressWarnings({ "resource", "unused" })
	private void initIgnoreList() {
		DataInputStream in = null;
		try {
			String executionPath = System.getProperty("user.dir");
			File ignore = new File(executionPath + File.separator
					+ "ztl.ignore");
			FileInputStream fstream = new FileInputStream(ignore);
			in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;

			String key = null;
			List<String> list = new ArrayList<String>();
			boolean commented = false;
			while ((strLine = br.readLine()) != null) {
				strLine = strLine.trim();
				if (commented) {
					if (strLine.startsWith("*/")) {
						commented = false;
						continue;
					}
					continue;
				}
				if (strLine.startsWith("/*")) {
					commented = true;
					continue;
				}
				if (strLine.isEmpty() || strLine.startsWith("#")
						|| strLine.startsWith("//"))
					continue;

				int keyIndex = strLine.indexOf("={");
				if (keyIndex > -1) {
					if (key != null) {
						throw new IllegalArgumentException(
								"The file format is wrong! No end '}' was found! ["
										+ key + "]");
					}
					key = strLine.substring(0, keyIndex);
					continue;
				}

				if (strLine.startsWith("}")) {
					if (key == null) {
						throw new IllegalArgumentException(
								"The file format is wrong! No key was found!");
					}
					String[] keys = key.split(",");
					for (String k : keys) {
						k = k.replaceAll("\\*", ".*");
						// System.err.println("put: " + k +"=" + list);
						if (_ignoreMap.containsKey(k)) {
							_ignoreMap.get(k).addAll(list);
						} else _ignoreMap.put(k, new ArrayList<String>(list));
					}
					list = new ArrayList<String>();
					key = null;
					continue;
				} else {
					// System.err.println("ignore: " + key +"=" + strLine);
					list.add(strLine);
				}
			}
		} catch (Exception e) {
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}

	private void addBrowserNameSetting(String browserName, String browserPath) {
		browserName = browserName.toLowerCase();
		String setting = browserPath;

		// for remote driver
		if (browserPath.indexOf(";") != -1) {
			String[] tokens = browserPath.split(";");
			_browserRemote.put(browserName, tokens[0]);
			if (tokens.length > 1) {
				browserPath = tokens[1];
			} else {
				browserPath = "";
			}
			setting = browserPath;
		}
		_driverSetting.put(browserName, setting);
	}

	private void initProperty() throws Exception, IOException {
		InputStream in = null;
		if (_properties == null) {
			try {
				String executionPath = System.getProperty("user.dir");
				
				_properties = new Properties();
				
				//load default properties
				File[] configFiles = new File[] {
						new File(executionPath + File.separator + CONFIG_NAME),
						new File(executionPath + File.separator + LOCAL_CONFIG_NAME)};
				InputStream is = null;
				for (File f: configFiles) {
					try {
						if (f.exists()) {
							is = new FileInputStream(f);
							_properties.load(is);
							System.out.println("Load config "+f.getAbsolutePath());
						}
					} catch(Exception x) {
						System.err.print(x.getMessage());
					} finally {
						if (is != null) {
							try {
								is.close();
							} catch (IOException e) {}
							is = null;
						}
					}
				}

				// _client = _properties.getProperty("client");
				_driverMap = new HashMap<String, WebDriver>(15);
				_debuggable = Boolean.parseBoolean(_properties
						.getProperty("debuggable"));
				_contextPath = _properties.getProperty("context-path");
				_delay = _properties.getProperty("delay");
				_browser = _properties.getProperty("browser");
				_timeout = _properties.getProperty("timeout");
				_zsstheme = _properties.getProperty("zsstheme");
				//_comparable = Boolean.parseBoolean(_properties
				//		.getProperty("comparable", "false"));
				//_granularity = Integer.parseInt(_properties
				//		.getProperty("granularity"));
				//_leniency = Integer.parseInt(_properties.getProperty("leniency"));
				for (Iterator<Map.Entry<Object, Object>> iter = _properties.entrySet().iterator();
						iter.hasNext();) {
					final Map.Entry<Object, Object> setting = (Map.Entry<Object, Object>)
							iter.next();
					final String key = (String) setting.getKey();
					if (isBrowserSetting(key)) {
						addBrowserNameSetting(key, (String) setting.getValue());
						continue;
					}
				}
				_server = System.getProperty("server");
				if (_server == null || _server.isEmpty() || "${server}".equals(_server)) {
					_server = _properties.getProperty("server");
				}
				String allBrowser = System.getProperty("browser");
				if (allBrowser == null || allBrowser.isEmpty() || "${browser}".equals(allBrowser)) {
					allBrowser = _properties.getProperty(ALL_BROWSERS);
				}
				String[] allBrowsers = allBrowser.split(",");
				for (String browser : allBrowsers) {
					final String browserKey = browser.trim();
					if (_driverSetting.containsKey(browserKey)) {
						_allBrowsers.add(browserKey);
					}
				}

				// System Properties
				String sysprop = _properties.getProperty("systemproperties");
				if (sysprop != null) {
					String[] keys = sysprop.split(";");
					for (String key : keys) {
						int start = key.indexOf(":");
						if (start < 0) {
							System.err.println("The syntax of the system property is wrong! ["
									+ key + "]");
							continue;
						}
						System.setProperty(key.substring(0, start), key.substring(start + 1));
					}
				}
			} finally {
				if (in != null) {
					in.close();
				}
			}
		}
	}

	private boolean isBrowserSetting(String str) {
		for (String browserStr : SUPPORTED_BROWSER) {
			if (str.toLowerCase().startsWith(browserStr))
				return true;
		}
		return false;
	}

	/**
	 * TODO Logging untested yet
	 */
	protected static final String RESULT_FILE_ENCODING = "UTF-8";

	protected static final String RESULTS_BASE_PATH = "Log";

	/**
	 * Returns whether in a comparable mode.
	 * <p>
	 * default: false. Property name in config.properties: <b>comparable</b>,
	 * which means the image comparison is in a comparable mode, if true
	 * specified. Otherwise, the image is stored to the imgsrc directory as the
	 * base images.
	 * 
	 */
	public boolean isComparable() {
		return _comparable;
	}

	/**
	 * Returns the granularity for each comparing section.
	 * <p>
	 * Property name in config.properties: <b>granularity</b>
	 * <p>
	 * It is better to have 1~15, less is a precise comparison, but performance
	 * is slow. Don't specify too high, it may compare without any different.
	 */
	public int getGranularity() {
		return _granularity;
	}

	/**
	 * Returns the leniency for each comparing section.
	 * <p>
	 * Property name in config.properties: <b>leniency</b>
	 * <p>
	 * It is better to have 1~10, less is a precise comparison.
	 */
	public int getLeniency() {
		return _leniency;
	}

}
