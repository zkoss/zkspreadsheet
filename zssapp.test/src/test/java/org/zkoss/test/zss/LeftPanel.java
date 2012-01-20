/* LeftPanel.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 3:35:52 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.Widget;

import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public class LeftPanel extends Widget {

	/**
	 * @param widgetScript
	 * @param webDriver
	 */
	/*package*/ LeftPanel(String widgetScript, JQueryFactory jqFactory, ConditionalTimeBlocker au, WebDriver webDriver) {
		super(widgetScript, jqFactory, au, webDriver);
	}
	
	public Header getRowHeader(int row) {
		String script = widgetScript() + ".getHeader(" + row + ")";
		return new Header(script, jqFactory, timeBlocker, webDriver);
	}
}
