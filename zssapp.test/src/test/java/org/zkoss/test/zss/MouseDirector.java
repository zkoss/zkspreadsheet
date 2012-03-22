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
import org.openqa.selenium.WebDriverException;
import org.zkoss.test.Browser;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.Keycode;
import org.zkoss.test.MouseButton;
import org.zkoss.test.Style;

import com.google.common.base.Predicate;
import com.google.inject.Inject;

/**
 * @author sam
 *
 */
public class MouseDirector {

	final WebDriver webDriver;
	
	final Spreadsheet spreadsheet;
	
	final Browser browser;
	
	final TopPanel topPanel;
	
	final ConditionalTimeBlocker timeBlocker;
	
	final JQueryFactory jqFactory;
	
	@Inject
	/*package*/ MouseDirector (Spreadsheet spreadsheet, TopPanel topPanel, JQueryFactory jqFactory,
			Browser browser, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		this.spreadsheet = spreadsheet;
		this.topPanel = topPanel;
		this.timeBlocker = timeBlocker;
		this.webDriver = webDriver;
		this.browser = browser;
		this.jqFactory = jqFactory;
	}
	
	public void openCellContextMenu(int row, int col) {
		openCellContextMenu(row, col, row, col);
	}
	
	public void openCellContextMenu(int tRow, int lCol, int bRow, int rCol) {
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		JQuery target = spreadsheet.getCell(tRow, lCol).jq$n();
		new JavascriptActions(webDriver)
			.contextClick(target)
			.perform();
	}
	
	public void rightClick(JQuery target) {
		try { //event may hide the element, cause next event throw exception
			new JavascriptActions(webDriver)
			.mouseOver(target, MouseButton.RIGHT)
			.mouseDown(target, MouseButton.RIGHT)
			.mouseUp(target, MouseButton.RIGHT)
			.contextClick(target)
			.perform();
		} catch (WebDriverException ex) {
		}
		
		if (browser.isSafari() || browser.isGecko()) {
			timeBlocker.waitUntil(1);
		}
		timeBlocker.waitResponse();
	}
	
	public void click(final JQuery target) {
		//TODO: test getWebElement().click() with selenium 2.20 version
		// doesn't work on IE8 with selenium 2.11.0 version
		// doesn't work well on FF7 with selenium 2.15.0
		//target.getWebElement().click();
		
		timeBlocker.waitUntil(new Predicate<Void>() {
			
			@Override
			public boolean apply(Void input) {
				return target.isVisible();
			}
		});
		try { //event may hide the element, cause next event throw exception
			new JavascriptActions(webDriver)
			.mouseOver(target, MouseButton.LEFT)
			.mouseDown(target, MouseButton.LEFT)
			.mouseUp(target, MouseButton.LEFT)
			.click(target)
			.perform();
		} catch (WebDriverException ex) {
		}
		
		if (browser.isSafari() || browser.isGecko()) {
			timeBlocker.waitUntil(1);
		}
		timeBlocker.waitResponse();
	}
	
	public void openRowContextMenu(int row) {
		openRowContextMenu(row, row);
	}
	
	public void openRowContextMenu(int tRow, int bRow) {
		selectRow(tRow, bRow);
		
		JQuery target = spreadsheet.getRow(tRow).jq$n();
		new JavascriptActions(webDriver)
			.mouseOver(target, MouseButton.RIGHT)
			.mouseDown(target, MouseButton.RIGHT)
			.mouseUp(target, MouseButton.RIGHT)
			.contextClick(target)
			.perform();
	}

	public void selectRow(int tRow, int bRow) {
		selectRow(tRow);
		
		JQuery from = spreadsheet.getRow(tRow).jq$n();
		JQuery to = spreadsheet.getRow(bRow).jq$n();		
		new JavascriptActions(webDriver)
			.mouseOver(from, MouseButton.LEFT)
			.mouseDown(from, MouseButton.LEFT)
			.mouseMove(from, MouseButton.LEFT)
			.mouseMove(to, MouseButton.LEFT)
			.mouseUp(to, MouseButton.LEFT)
			.perform();
	}

	public void selectRow(int tRow) {
		JQuery c = spreadsheet.getRow(tRow).jq$n();
		new JavascriptActions(webDriver)
			.mouseOver(c, MouseButton.LEFT)
			.mouseDown(c, MouseButton.LEFT)
			.mouseUp(c, MouseButton.LEFT)
			.click(c)
			.perform();
	}

	public void openColumnContextMenu(int col) {
		openColumnContextMenu(col, col);
	}
	
	public void openColumnContextMenu(int lCol, int rCol) {
		selectColumn(lCol, rCol);
		
		JQuery target = spreadsheet.getColumn(lCol).jq$n();
		new JavascriptActions(webDriver)
			.mouseOver(target, MouseButton.RIGHT)
			.mouseDown(target, MouseButton.RIGHT)
			.mouseUp(target, MouseButton.RIGHT)
			.contextClick(target)
			.perform();
	}
	
	public void selectColumn(int col) {
		JQuery c = spreadsheet.getColumn(col).jq$n();
		new JavascriptActions(webDriver)
			.mouseOver(c, MouseButton.LEFT)
			.mouseDown(c, MouseButton.LEFT)
			.mouseUp(c, MouseButton.LEFT)
			.click(c)
			.perform();
	}
	
	public void selectColumn(int lCol, int rCol) {
		//click
		selectColumn(lCol);
		
		JQuery from = spreadsheet.getColumn(lCol).jq$n();
		JQuery to = spreadsheet.getColumn(rCol).jq$n();		
		new JavascriptActions(webDriver)
			.mouseOver(from, MouseButton.LEFT)
			.mouseDown(from, MouseButton.LEFT)
			.mouseMove(from, MouseButton.LEFT)
			.mouseMove(to, MouseButton.LEFT)
			.mouseUp(to, MouseButton.LEFT)
			.perform();
	}
	
	public void dragColumnToHide(int col) {
		dragColumnToResize(col, 0);
	}
	
	public void dragColumnToResize(int col, int newSize) {
		Header columnHeader = topPanel.getColumnHeader(col);
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

	public void dragEdit(int srcRow, int srcCol, int targetRow, int targetCol) {
		dragEdit(srcRow, srcCol, srcRow, srcCol, targetRow, targetCol);
	}
	
	public void dragEdit(int srcTopRow, int srcLeftCol, int srcBtmRow, int srcRightCol, 
			int targetRow, int targetCol) {
		spreadsheet.setSelection(srcTopRow, srcLeftCol, srcBtmRow, srcRightCol);
		
		JQuery drag = jqFactory.create("'.zsselect'");
		JQuery dropAt = spreadsheet.getCell(targetRow, targetCol).jq$n();
		dragAction(drag, dropAt);
	}
	
	public void dragFill(int srcRow, int srcCol, int targetRow, int targetCol) {
		dragFill(srcRow, srcCol, srcRow, srcCol, targetRow, targetCol);
	}
	
	public void dragFill(int srcTopRow, int srcLeftCol, int srcBtmRow, int srcRightCol, 
			int targetRow, int targetCol) {
		spreadsheet.setSelection(srcTopRow, srcLeftCol, srcBtmRow, srcRightCol);
		
		JQuery drag = jqFactory.create("'.zsseldot'");
		JQuery dropAt = spreadsheet.getCell(targetRow, targetCol).jq$n();
		dragAction(drag, dropAt);
	}
	
	private void dragAction(JQuery drag, JQuery dropAt) {
		new JavascriptActions(webDriver)
		.mouseOver(drag, MouseButton.LEFT)
		.mouseDown(drag, MouseButton.LEFT)
		.mouseMove(drag, MouseButton.LEFT)
		.mouseMove(dropAt, MouseButton.LEFT)
		.mouseUp(dropAt, MouseButton.LEFT)
		.perform();
		
		timeBlocker.waitResponse();
	}
	
	public void pageUp(int row, int col) {
		spreadsheet.focus(row, col);
		
		JQuery target = spreadsheet.jq$n();
		int keycode = Keycode.PAGE_UP.intValue();
		new JavascriptActions(webDriver)
		.keyDown(target, keycode)
		.keyUp(target, keycode)
		.perform();
		
		timeBlocker.waitResponse();
	}

	public void pageDown(int row, int col) {
		spreadsheet.focus(row, col);
		
		JQuery target = spreadsheet.jq$n();
		int keycode = Keycode.PAGE_DOWN.intValue();
		new JavascriptActions(webDriver)
		.keyDown(target, keycode)
		.keyUp(target, keycode)
		.perform();
		
		timeBlocker.waitResponse();
	}
}
