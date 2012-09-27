/* Browser.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 20, 2012 3:25:04 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.util.Properties;


/**
 * @author sam
 *
 */
public class Config {
	
	public enum Browser {
		FIREFOX, CHROME, IE, OPERA;
	}
	
//	@Inject
//	public Config (String configFilePath) {
//	}
	
	private Properties setting;
	public void init() {
		//1. try to load .prop file if exist
		//2. if doesn't exist, write default setting to .prop file
	}
	
	//TODO: get value from config file
	public Browser getBrowser() {
//		if (setting == null) {
//			throw new IllegalStateException("Config shall invoke Config.init() first");
//		}
//		return Browser.FIREFOX;
		return Browser.CHROME;
//		return Browser.IE;
	}
}
