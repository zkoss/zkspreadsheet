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
import com.google.inject.name.Named;


/**
 * @author sam
 *
 */
public class Spreadsheet extends Widget {
	
	private enum SpreadsheetField {
		SheetCtrl("sheetCtrl");
		
		private String field;
		SpreadsheetField (String field) {
			this.field = field;
		}
		
		@Override
		public String toString() {
			return field;
		}
	}
	
	@Inject
	public Spreadsheet (@Named("Spreadsheet Id") String spreadsheetId, JQueryFactory jqFactory, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		super(new Id(spreadsheetId), jqFactory, timeBlocker, webDriver);
	}
	
	private String sheetCtrlScript() {
		return widgetScript() + "." + SpreadsheetField.SheetCtrl;
	}
	
	//TODO: move to Cell class
	private String cellScript(int row, int col) {
		return widgetScript() + "." + SpreadsheetField.SheetCtrl + ".getCell(" + row + "," + col + ")";
	}
	
	//TODO: move to Cell class
	private String cellDOMElementScript(int row, int col) {
		return cellScript(row, col) + ".$n()";
	}
	
	//TODO: move to Cell class
	private String returnCellScript(int row, int col) {
		return "return " + cellScript(row, col);
	}
	
	public String getCellEdit(int row, int col) {
		String script = returnCellScript(row, col) + ".text";
		return (String) javascriptExecutor.executeScript(script);
	}

	/**
	 * @param row
	 * @param col
	 */
	public void focus(int row, int col) {
//		WebElement cell = getCell(col, row).getWebElement();
//		cell.click();
//		cell.click();
		
		JQuery target = jqFactory.create(cellDOMElementScript(row, col));
		
		//if spreadsheet widget doesn't have focus, the first event will focus on last focus
		new JavascriptActions(webDriver)
		.mouseDown(target, MouseButton.LEFT)
		.mouseUp(target, MouseButton.LEFT).perform();
		
		new JavascriptActions(webDriver)
		.mouseDown(target, MouseButton.LEFT)
		.mouseUp(target, MouseButton.LEFT).perform();
		timeBlocker.waitResponse();
	}
	
	//TODO
	public void select(int tRow, int lCol, int bRow, int rCol) {
		
	}
	
	/**
	 * Returns the focus representation of spreadsheet widget
	 * 
	 * @return
	 */
	public JQuery jq$focus() {
		String script = widgetScript() + ".$n('fo')";
		return jqFactory.create(script);
	}
	
	//TODO: use injection
	public InlineEditor getInlineEditor() {
		String script = sheetCtrlScript() + ".inlineEditor";
		return new InlineEditor(script, jqFactory, timeBlocker, webDriver);
	}
	
	//TODO: use injection
	public FormulabarEditor getFormulabarEditor() {
		String script = sheetCtrlScript() + ".formulabarEditor";
		return new FormulabarEditor(script, jqFactory, timeBlocker, webDriver);
	}
	
	public TopPanel getTopPanel() {
		String script = sheetCtrlScript() + ".tp";
		return new TopPanel(script, jqFactory, timeBlocker, webDriver);
	}
	
	public LeftPanel getLeftPanel() {
		String script = sheetCtrlScript() + ".lp";
		return new LeftPanel(script, jqFactory, timeBlocker, webDriver);
	}
	
	public Row getRow(int row) {
		return getMainBlock().getRow(row);
	}
	
	public MainBlock getMainBlock() {
		String script = sheetCtrlScript() + ".activeBlock";
		return new MainBlock(script, jqFactory, timeBlocker, webDriver);
	}
	
	//TODO: 
//	public DOMElement getCellDOMElement(int row, int col) {
//		return null;
//	}
}
