/* Browser.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 10:38:53 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.lang.reflect.Method;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;

import com.google.inject.Inject;

/**
 * @author sam
 *
 */
public class Browser {
	
	private final JavascriptExecutor javascriptExecutor;

	@Inject
	/*package*/ Browser(WebDriver webDriver) {
		javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	public boolean isSafari() {
		Object obj = javascriptExecutor.executeScript("return zk.safari");
		return obj != null;
	}
	
	public boolean isGecko() {
		return javascriptExecutor.executeScript("return zk.gecko") != null;
	}
	
	public boolean isIE() {
    	Object obj =  javascriptExecutor.executeScript("return zk.ie");
    	Integer verNum = null;
		Method m = null;
		try {
			if (obj != null) {
				m = obj.getClass().getMethod("intValue");
				verNum = (Integer)m.invoke(obj);
			}
		} catch (Exception ex) {
			throw new RuntimeException("can't find intValue method");
		}
    	return verNum != null && verNum >= 6;
    }
	
	public boolean isIE(int version) {
    	Object obj =  javascriptExecutor.executeScript("return zk.ie");
    	Integer ver = null;
    	Method m = null;
		try {
			if (obj != null) {
				m = obj.getClass().getMethod("intValue");
				ver = (Integer)m.invoke(obj);
			}
		} catch (Exception ex) {
			throw new RuntimeException("can't find intValue method");
		}
    	return ver != null && ver == version;
    }
	
	public boolean isIE7() {
    	return isIE() && isIE(7);
    }
	
	public boolean isIE6() {
    	return isIE() && isIE(6);
    }
}
