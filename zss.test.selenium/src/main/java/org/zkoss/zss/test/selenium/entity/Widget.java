/* Widget.java

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
 * A simulator of the ZK client widget.
 * @author RaymondChao
 *
 */
public class Widget extends ClientWidget {
	
	private static String WIDGET_ELEMENT = "zk.Widget.$(%1$s)";
	
	private static String RETURN_$N = "return %1$s.$n('%2$s');";
	
	protected Widget() {
		super();
	}
	
	protected Widget(JQuery jQuery) {
		this();
		setSelector(String.format(WIDGET_ELEMENT, jQuery.toString()));
	}
	
	public static Widget $(JQuery jQuery) {
		if (jQuery == null) {
			throw new NullPointerException("Selector cannot be null.");
		}
		return new Widget(jQuery);
	}
	
	public WebElement $n() {
		return $n("");
	}
	
	public WebElement $n(String uuid) {
		WebElement element = TestCaseBase.driver().findElement(ZSSBy.javascript(
				String.format(RETURN_$N, toString(), uuid)));
		if (element == null) {
			throw new RuntimeException("can't find element.");
		}
		return element;
	}
	
	public Widget getChild(String name) {
		Widget child = new Widget();
		child.setSelector(this.getSelector() + "." + toUpperCase("get", name) + "()");
		return child;
	}
	
	public WebElement toWebElement() {
		return $n();
	}
	
	public void focus() {
		call("focus");
		TestCaseBase.eval(this);
	}
}
