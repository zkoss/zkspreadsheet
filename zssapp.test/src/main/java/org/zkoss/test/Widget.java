/* Widget.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 4:16:33 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import javax.annotation.Nullable;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public abstract class Widget {
	
	private final static String WIDGET = "zk.Widget.$('$%1')";
	
	private final String widgetScript;
	
	protected final ConditionalTimeBlocker timeBlocker;
	
	protected final WebDriver webDriver;
	
	protected final JavascriptExecutor javascriptExecutor;
	
	protected final JQueryFactory jqFactory;
	
	@Inject
	public Widget (@Assisted Id widgetId, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		widgetScript = WIDGET.replace("%1", widgetId.toString());
		this.timeBlocker = timeBlocker;
		this.jqFactory = jqFactory;
		this.webDriver = webDriver;
		javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	@Inject
	public Widget (@Assisted String widgetScript, JQueryFactory jqFactory, ConditionalTimeBlocker au, WebDriver webDriver) {
		this.widgetScript = Preconditions.checkNotNull(widgetScript);
		this.jqFactory = jqFactory;
		this.timeBlocker = au;
		this.webDriver = webDriver;
		javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	public String widgetScript () {
		return widgetScript;
	}
	
	/**
	 * Returns the {@link JQuery} selector that represent 
	 * the child element of the DOM element(s) that this widget is bound to
	 */
	public String $n() {
		return widgetScript + ".$n()";
	}
	
	/**
	 * Returns the {@link JQuery} selector that represent 
	 * the child element of the DOM element(s) that this widget is bound to
	 */	
	public String $n(@Nullable String subId) {
		if (Strings.isNullOrEmpty(subId)) {
			return $n();
		}
		return widgetScript + ".$n('" + subId + "')";
	}
		
	/**
	 * Returns the {@link JQuery} of the DOM element that this widget is bound to
	 * @return
	 */
	public JQuery jq$n() {
		return jqFactory.create($n());
	}
	
	/**
	 * Returns the {@link JQuery} selector of the widget
	 * 
	 * @return String
	 */
//	private String getSelector() {
//		return "'$" + id + "'";
//	}
	
	public boolean isVisible() {
		return jq$n().isVisible();
	}
	
	/**
	 * Returns the width of this widget.
	 * 
	 * @return
	 */
	public int getWidth() {
		return jq$n().width();
	}
	
	/**
	 * Returns the height of this widget.
	 * 
	 * @return
	 */
	public int getHeight() {
		return jq$n().height();
	}
	
	/**
	 * Returns the {@link WebElement} of the widget that bound to
	 * 
	 * @return
	 */
	public WebElement getWebElement() {
		return jq$n().getWebElement();
	}
}
