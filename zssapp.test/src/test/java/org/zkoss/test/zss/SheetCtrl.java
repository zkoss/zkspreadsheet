/* SheetCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 9:31:53 AM , Created by sam
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

/**
 * @author sam
 *
 */
public class SheetCtrl extends Widget {

	/**
	 * @param widgetScript
	 * @param jqFactory
	 * @param au
	 * @param webDriver
	 */
	@Inject
	public SheetCtrl(Spreadsheet spreadsheet, JQueryFactory jqFactory,
			ConditionalTimeBlocker au, WebDriver webDriver) {
		super(spreadsheet.widgetScript() + ".sheetCtrl", jqFactory, au, webDriver);
	}
	
	public Rect getLastSelection() {
		String lastSelectionScript = "return " + widgetScript() + ".getLastSelection()";
		String hasSelectionScript = lastSelectionScript + " != null";
		
		Boolean hasSelection = (Boolean)javascriptExecutor.executeScript(hasSelectionScript);
		if (hasSelection == null || !hasSelection) {
			return null;
		}
		
		Object lCol = javascriptExecutor.executeScript(lastSelectionScript + ".left");
		Object rCol = javascriptExecutor.executeScript(lastSelectionScript + ".right");
		Object tRow = javascriptExecutor.executeScript(lastSelectionScript + ".top");
		Object bRow = javascriptExecutor.executeScript(lastSelectionScript + ".bottom");
			
		return new Rect(Util.intValue(lCol), Util.intValue(tRow), Util.intValue(rCol), Util.intValue(bRow));
	}

	public boolean isHighlightVisible() {
		return jqFactory.create("'.zshighlight'").isVisible();
	}
	
	/**
	 * @return
	 */
	public Rect getLastHighlight() {
		
		String lastHighlightScript = "return " + widgetScript() + ".getLastHighlight()";
		String hasHighlightScript = lastHighlightScript + " != null";
		Boolean hasHighlight = (Boolean)javascriptExecutor.executeScript(hasHighlightScript);
		if (hasHighlight == null || !hasHighlight) {
			return null;
		}
		
		Object lCol = javascriptExecutor.executeScript(lastHighlightScript + ".left");
		Object rCol = javascriptExecutor.executeScript(lastHighlightScript + ".right");
		Object tRow = javascriptExecutor.executeScript(lastHighlightScript + ".top");
		Object bRow = javascriptExecutor.executeScript(lastHighlightScript + ".bottom");
		
		return new Rect(Util.intValue(lCol), Util.intValue(tRow), Util.intValue(rCol), Util.intValue(bRow));
	}
}
