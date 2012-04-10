/* Spreadsheet.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 12:21:44 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.Id;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.MouseButton;
import org.zkoss.test.Widget;

import com.google.inject.Inject;
import com.google.inject.Injector;
import com.google.inject.name.Named;


/**
 * @author sam
 *
 */
public class Spreadsheet extends Widget {

	private final Cell.Factory cellFactory;
	private final Injector injector;
	
	@Inject
	public Spreadsheet (@Named("Spreadsheet Id") String spreadsheetId, Injector injector, 
			Cell.Factory cellFactory, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		super(new Id(spreadsheetId), jqFactory, timeBlocker, webDriver);
		
		this.cellFactory = cellFactory;
		this.injector = injector;
	}

	/**
	 * @param row
	 * @param col
	 */
	public void focus(int row, int col) {
//		WebElement cell = getCell(col, row).getWebElement();
//		cell.click();
//		cell.click();
		
		JQuery target = cellFactory.create(row, col).jq$n();
		
		//if spreadsheet widget doesn't have focus, the first event will focus on last focus
		new JavascriptActions(webDriver)
		.mouseDown(target, MouseButton.LEFT)
		.mouseUp(target, MouseButton.LEFT)
		.perform();
		
		new JavascriptActions(webDriver)
		.mouseDown(target, MouseButton.LEFT)
		.mouseUp(target, MouseButton.LEFT)
		.perform();
		timeBlocker.waitResponse();
	}
	
	public void setSelection(int tRow, int lCol, int bRow, int rCol) {
		focus(tRow, lCol);
		
		JQuery from = cellFactory.create(tRow, lCol).jq$n();
		JQuery to = cellFactory.create(bRow, rCol).jq$n();
		
		new JavascriptActions(webDriver)
		.mouseDown(from, MouseButton.LEFT)
		.mouseMove(from, MouseButton.LEFT)
		.mouseMove(to, MouseButton.LEFT)
		.mouseUp(to, MouseButton.LEFT)
		.perform();
		timeBlocker.waitResponse();
	}
	
	public boolean isSelection(int row, int col) {
		return isSelection(row, col, row, col);
	}
	
	public boolean isSelection(int tRow, int lCol, int bRow, int rCol) {
		SheetCtrl sheet = injector.getInstance(SheetCtrl.class);
		
		Rect selection = sheet.getLastSelection();
		if (selection == null) {
			return false;
		}
		
		return selection.getTop() == tRow
			&& selection.getLeft() == lCol
			&& selection.getBottom() == bRow
			&& selection.getRight() == rCol;
	}
	
	public Rect getVisibleRange() {
		SheetCtrl sheet = injector.getInstance(SheetCtrl.class);
		return sheet.getVisibleRange();
	}
	
	public boolean isHighlight(int tRow, int lCol, int bRow, int rCol) {
		SheetCtrl sheet = injector.getInstance(SheetCtrl.class);
		if (!sheet.isHighlightVisible())
			return false;
		
		Rect highlight = sheet.getLastHighlight();
		if (highlight == null) {
			return false;
		}
		
		return highlight.getTop() == tRow
			&& highlight.getLeft() == lCol
			&& highlight.getBottom() == bRow
			&& highlight.getRight() == rCol;
	}
	
	protected Cell getCell(int row, int col) {
		return cellFactory.create(Integer.valueOf(row), Integer.valueOf(col));
	}
	
	public JQuery jq$focus() {
		return getSheetCtrl().jq$n("fo");
	}
	
	public InlineEditor getInlineEditor() {
		return injector.getInstance(InlineEditor.class);
	}
	
	public FormulabarEditor getFormulabarEditor() {
		return injector.getInstance(FormulabarEditor.class);
	}
	
	public SheetCtrl getSheetCtrl() {
		return injector.getInstance(SheetCtrl.class);
	}
	
	public TopPanel getTopPanel() {
		return injector.getInstance(TopPanel.class);
	}
	
	public LeftPanel getLeftPanel() {
		return injector.getInstance(LeftPanel.class);
	}
	
	public Header getRowHeader(int row) {
		return injector.getInstance(LeftPanel.class).getRowHeader(row);
	}
	
	public MainBlock getMainBlock() {
		return injector.getInstance(MainBlock.class);
	}
	
	public Header getColumnHeader(int col) {
		return injector.getInstance(TopPanel.class).getColumnHeader(col);
	}
}
