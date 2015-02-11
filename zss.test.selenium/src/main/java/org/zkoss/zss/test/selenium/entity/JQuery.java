/* JQuery.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Fri, Jul 11, 2014  5:05:17 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.TestCaseBase;

/**
 * A simulator of jQuery client side object, which wraps the JQuery client side
 * API.
 *
 * @author jumperchen
 * @author RaymondChao
 *
 */
public class JQuery extends ClientWidget {

	/**
	 * The script of get jq by selector
	 */
	private static final String JQ = "jq('%1$s')";
	private static String PROP_RESULT = "return %1$s.%2$s;";
	
	private JQuery() {
		super();
	}
	
	private JQuery(String selector) {
		super();
		setSelector(String.format(JQ, selector));
	}
	
	public static JQuery $(String selector) {
		if (isEmpty(selector)) {
			throw new NullPointerException("selector cannot be null.");
		}
		return new JQuery(selector);
	}
	
	public String text() {
		return (String) fun("text");
	}
	
	public String html() {
		return (String) fun("html");
	}
	
	public String val() {
		return (String) fun("val");
	}
	
	public String css(String propertyName) {
		return TestCaseBase.eval("return " + toString() + ".css('" + propertyName + "')").toString().trim();
	}
	
	public int width() {
		return Integer.parseInt(fun("width").toString());
	}
	
	public int height() {
		return Integer.parseInt(fun("height").toString());
	}
	
	public int offsetLeft() {
		return (int) Math.round(Double.parseDouble(super.getProperty("offset().left").toString()));
	}
	
	public int offsetTop() {
		return (int) Math.round(Double.parseDouble(super.getProperty("offset().top").toString()));
	}
	
	public boolean exists() {
		return length() > 0;
	}
	
	public int length() {
		return ((Long) prop("length")).intValue();
	}
	
	public Object prop(String name) {
		return TestCaseBase.eval(String.format(PROP_RESULT, toString(), name));
	}
	
	public Object getProperty(String name) {
		return prop(name);
	}
	
	/**
	 * proxy for jQuery get method
	 * @param index
	 * @return Element  the DOM element
	 */
	public WebElement get(int index) {
		WebElement element = TestCaseBase.driver().findElement(ZSSBy.javascript("return " + toString() + "[" + index + "]"));
		if (element == null) {
			throw new RuntimeException("can't find element.");
		}
		return element;
	}
	
	public WebElement toWebElement() {
		return get(0);
	}
}
