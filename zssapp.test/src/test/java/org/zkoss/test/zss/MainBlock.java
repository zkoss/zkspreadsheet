/* MainBlock.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 1, 2012 2:44:12 PM , Created by sam
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
public class MainBlock extends Widget {

	/*package*/ MainBlock(String widgetScript, JQueryFactory jqFactory,
			ConditionalTimeBlocker au, WebDriver webDriver) {
		super(widgetScript, jqFactory, au, webDriver);
	}
	
	public Row getRow(int row) {
		String script = widgetScript() + ".getRow(" + row + ")";
		return new Row(script, jqFactory, timeBlocker, webDriver);
	}
}
