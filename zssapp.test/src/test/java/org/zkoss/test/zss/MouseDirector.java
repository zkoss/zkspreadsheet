/* MouseDirector.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 4:10:17 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.MouseButton;
import org.zkoss.test.Style;

import com.google.inject.Inject;

/**
 * @author sam
 *
 */
public class MouseDirector {

	final WebDriver webDriver;
	
	final Spreadsheet spreadsheet;
	
	final ConditionalTimeBlocker timeBlocker;
	
	@Inject
	/*package*/ MouseDirector (Spreadsheet spreadsheet, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		this.spreadsheet = spreadsheet;
		this.timeBlocker = timeBlocker;
		this.webDriver = webDriver;
	}
	
	public void dragColumnToHide(int col) {
		dragColumnToResize(col, 0);
	}
	
	public void dragColumnToResize(int col, int newSize) {
		Header columnHeader = spreadsheet.getTopPanel().getColumnHeader(col);
		JQuery from = columnHeader.getBoundary();
		
		int xOffset = newSize - columnHeader.getWidth();
		new JavascriptActions(webDriver)
		.mouseOver(from, MouseButton.LEFT)
		.mouseDown(from, MouseButton.LEFT)
		.mouseMove(from, MouseButton.LEFT)
		.mouseMove(from, MouseButton.LEFT, xOffset, 0)
		.mouseUp(from, MouseButton.LEFT)
		.perform();
		
		timeBlocker.waitResponse();
	}
	
	public void dragRowToHide(int row) {
		dragRowToResize(row, 0);
	}
	
	public void dragRowToResize(int row, int newSize) {
		Header rowHeader = spreadsheet.getLeftPanel().getRowHeader(row);
		JQuery from = rowHeader.getBoundary();
		
		int headerBorder = rowHeader.jq$n().zk().sumStyles("tb", Style.BORDERS);
		int yOffset = (newSize + headerBorder) - rowHeader.getHeight();
		new JavascriptActions(webDriver)
		.mouseOver(from, MouseButton.LEFT)
		.mouseDown(from, MouseButton.LEFT)
		.mouseMove(from, MouseButton.LEFT)
		.mouseMove(from, MouseButton.LEFT, 0, yOffset)
		.mouseUp(from, MouseButton.LEFT)
		.perform();
		
		timeBlocker.waitResponse();
	}
}
