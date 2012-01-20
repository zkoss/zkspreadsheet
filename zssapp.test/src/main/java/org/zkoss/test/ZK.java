/* ZK.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 4:52:39 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.util.List;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;

import com.google.common.base.Strings;

/**
 * @author sam
 *
 */
public class ZK {
	private String ZK = "return zk('%1')";
	protected StringBuffer _out;

	protected final WebDriver _webDriver;
	
	protected final JavascriptExecutor _javascriptExecutor;
	
	public ZK(String selector, WebDriver webDriver) {
		if (Strings.isNullOrEmpty(selector))
			throw new IllegalArgumentException("ZK selector cannot be null!");
		_out = new StringBuffer(ZK.replace("%1", selector));
		
		_webDriver = webDriver;
		_javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	public ZK(StringBuffer out, String script, WebDriver webDriver) {
		_out = new StringBuffer(out).append(script);
		
		_webDriver = webDriver;
		_javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	public String revisedOffsetScript() {
		String script = _out.toString() + ".revisedOffset();";
		return script.replace("return ", "");
	}
	
	public int[] revisedOffset() {
		List<Long> offset = (List)_javascriptExecutor.executeScript(_out.toString() + ".revisedOffset()");
		return new int[] {offset.get(0).intValue(), offset.get(1).intValue()};
	}
	
	/**
	 * Returns the summation of the specified styles.
	 * 
	 * @param String areas the areas is abbreviation for left "l", right "r", top "t", and bottom "b". So you can specify to be "lr" or "tb" or more. 
	 * @param Style, {@link Style#PADDINGS}, {@link Style#MARGINS} or {@link Style#BORDERS}.
	 * @return int the summation
	 */
	public int sumStyles(String areas, Style styles) {
		String script = _out.toString() + ".sumStyles('" +areas + "', jq." + styles + ")";
		//Firefox return Long
		//TODO: check other browsers, may return different type
		Long sum = (Long) _javascriptExecutor.executeScript(script);
		return sum.intValue();
	}
}
