/* SheetCtrl.java

	Purpose:
		
	Description:
		
	History:
		Mon, Jul 21, 2014 11:00:22 AM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

import org.zkoss.zss.test.selenium.entity.EditorWidget.EditorType;

/**
 * 
 * @author RaymondChao
 */
public class SheetCtrlWidget extends Widget {
	public SheetCtrlWidget(SpreadsheetWidget spreadsheet) {
		setSelector(spreadsheet.getSelector().toString() + ".sheetCtrl");
	}
	
	/**
	 * 
	 * @param row start from 0
	 * @param column row start from 0
	 * @return
	 */
	public CellWidget getCell(int row, int column) {
		return new CellWidget(this, row, column);
	}
	
	public CellWidget getCell(String cellRef) {
		return new CellWidget(this, cellRef);
	}

	public EditorWidget getInlineEditor() {
		return new EditorWidget(this);
	}
	public EditorWidget getFormulabarEditor() {
		return new EditorWidget(this, EditorType.FORMULABAR);
	}
	
	public Widget getContextMenu() {
		return this.getChild("styleMenupopup");
	}
}
