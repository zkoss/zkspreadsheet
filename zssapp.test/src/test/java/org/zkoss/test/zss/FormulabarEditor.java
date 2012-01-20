/* FormulabarEditor.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 9:58:35 AM , Created by sam
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

/**
 * @author sam
 *
 */
public class FormulabarEditor extends Widget {

	/**
	 * @param widgetId
	 * @param webDriver
	 */
	/*package*/ FormulabarEditor(String widgetScript, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		super(widgetScript, jqFactory, timeBlocker, webDriver);
	}
}
