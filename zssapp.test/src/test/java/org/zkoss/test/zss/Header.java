/* Header.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 3:31:44 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.Style;
import org.zkoss.test.Widget;
import org.zkoss.test.ZK;

/**
 * @author sam
 *
 */
public class Header extends Widget {

	/**
	 * @param widgetId
	 * @param webDriver
	 */
	/*package*/ Header(String widgetScript, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		super(widgetScript, jqFactory, timeBlocker, webDriver);
	}
	
	public JQuery getBoundary() {
		String script = $n("boun") + ".firstChild";
		return jqFactory.create(script);
	}

	/**
	 * Returns the header height (border + padding)
	 */
	public int getHeight() {
		if (!isVisible())
			return 0;
		
		int height = super.getHeight();
		final ZK zk = jq$n().zk();
		int extraHeight = zk.sumStyles("tb", Style.PADDINGS) + zk.sumStyles("tb", Style.BORDERS);
		return height + extraHeight;
	}
	
	/**
	 * Returns the header width (border + padding)
	 */
	@Override
	public int getWidth() {
		if (!isVisible())
			return 0;
		
		int width = super.getWidth();
		final ZK zk = jq$n().zk();
		int extraWidth = zk.sumStyles("lr", Style.PADDINGS) + zk.sumStyles("lr", Style.BORDERS);
		return width + extraWidth;
	}
}
