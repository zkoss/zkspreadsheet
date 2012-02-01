/* JQuery.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 12:39:39 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.lang.reflect.Method;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.google.common.base.Predicate;
import com.google.common.base.Strings;
import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public class JQuery {
	
	private final static String JQ = "return jq(%1)";
	
	private StringBuffer out;
	
	protected final ConditionalTimeBlocker timeBlocker;
	
	protected final JQueryFactory jqFactory;
	
	protected final WebDriver webDriver;
	
	protected final JavascriptExecutor javascriptExecutor;
	
	@Inject
	public JQuery(@Assisted String selector, ConditionalTimeBlocker timeBlocker, JQueryFactory jqFactory, WebDriver webDriver) {
		if (Strings.isNullOrEmpty(selector))
			throw new IllegalArgumentException("JQuery selector can not be empty");
		out = new StringBuffer(JQ.replace("%1", selector));
		
		this.timeBlocker = timeBlocker;
		this.jqFactory = jqFactory;
		this.webDriver = webDriver;
		javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	public String elementScript() {
		String script = out.toString() + "[0]";
		return script.replace("return ", "");
	}
	
	public ZK zk() {
		return new ZK(out, ".zk", webDriver);
	}
	
	public WebElement getWebElement() {
		return (WebElement)javascriptExecutor.executeScript(out.toString() + "[0]");
	}
	
	public boolean exists() {
		WebElement e = getWebElement();
		return e != null;
	}
	
	public boolean is(String selector) {
		return (Boolean)javascriptExecutor.executeScript(out.toString() + ".is('" + selector + "')");
	}
	
	public String css(String propertyName) {
		Object obj = javascriptExecutor.executeScript(out.toString() + ".css('" + propertyName + "')");
		return obj != null ? obj.toString() : "";
	}
	
	private boolean isVisibleImpl() {
		boolean isVisible = is(":visible");
		boolean display = !"hidden".equals(css("visibility"));
		return isVisible && display;
	}
	
	public boolean isVisible() {
		try {
			timeBlocker.waitUntil(1, new Predicate<Void>() {

				public boolean apply(Void input) {
					return exists() && isVisibleImpl();
				}
			});
		} catch (TimeoutException ex) {}
		return exists() && isVisibleImpl();
	}
	
	public int width() {
		Object width = javascriptExecutor.executeScript(out.toString() + ".width()");
		Method m = null;
		try {
			m = width.getClass().getMethod("intValue");
			return (Integer)m.invoke(width);
		} catch (Exception ex) {
			throw new RuntimeException("can't find intValue method");
		}
	}
	
	public int height() {
		Object h = javascriptExecutor.executeScript(out.toString() + ".height()");
		Method m = null;
		try {
			m = h.getClass().getMethod("intValue");
			return (Integer)m.invoke(h);
		} catch (Exception ex) {
			throw new RuntimeException("can't find intValue method");
		}
	}
	
	public int length() {
		Object length = javascriptExecutor.executeScript(out.toString() + ".length");
		Method m = null;
		try {
			m = length.getClass().getMethod("intValue");
			return (Integer)m.invoke(length);
		} catch (Exception ex) {
			throw new RuntimeException("can't find intValue method");
		}
	}
	
	public JQuery children() {
		return jqFactory.create(out.toString().replace("return ", "") + ".children()");
	}
	
	public JQuery children(String selector) {
		return jqFactory.create(out.toString().replace("return ", "") + ".children('" + selector +"')");
	}
	
	public JQuery first() {
		return jqFactory.create(out.toString().replace("return ", "") + ".first()");
	}
//	
//	public JQuery last() {
//		return new JQuery(out, ".last()", au, webDriver);
//	}
//	
//	public JQuery next() {
//		return new JQuery(out, ".next()", au, webDriver);
//	}
//	
//	public JQuery find(String arg) {
//		return new JQuery(out, ".find(\"" + arg + "\")", au, webDriver);
//	}
}
