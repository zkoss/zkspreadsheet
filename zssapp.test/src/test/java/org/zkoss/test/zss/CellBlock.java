/* CellBlock.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 1:57:09 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.Util;
import org.zkoss.test.Widget;

import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public class CellBlock extends Widget {

	/**
	 * @param widgetScript
	 * @param jqFactory
	 * @param au
	 * @param webDriver
	 */
	@Inject
	public CellBlock(@Assisted String widgetScript, JQueryFactory jqFactory,
			ConditionalTimeBlocker au, WebDriver webDriver) {
		super(widgetScript, jqFactory, au, webDriver);
	}
	
	public int getRowSize() {
		String script = "return " + widgetScript() + ".rows.length";
		Integer size = Util.intValue(javascriptExecutor.executeScript(script));
		if (size == null) {
			return 0;
		}
		return size;
	}
	
	public int getColumnSize() {
		String script = "return " + widgetScript() + ".rows[0].cells.length";
		Integer size = Util.intValue(javascriptExecutor.executeScript(script));
		if (size == null) {
			return 0;
		}
		return size;
	}
	
	public static interface Factory {
		public CellBlock create(String widgetScript);
	}
}