/* TopPanel.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 3:35:23 PM , Created by sam
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
public class TopPanel extends Widget {

	/**
	 * @param widgetScript
	 * @param webDriver
	 */
	/*package*/ TopPanel(String widgetScript, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		super(widgetScript, jqFactory, timeBlocker, webDriver);
	}

	public Header getColumnHeader(int col) {
		String script = widgetScript() + ".getHeader(" + col + ")";
		return new Header(script, jqFactory, timeBlocker, webDriver);
	}
}
