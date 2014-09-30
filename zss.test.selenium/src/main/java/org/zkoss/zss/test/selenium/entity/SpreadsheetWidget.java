/* ClientSpreadsheet.java

	Purpose:
		
	Description:
		
	History:
		Fri, Jul 11, 2014  6:56:24 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

/**
 * 
 * @author RaymondChao
 */
public class SpreadsheetWidget extends Widget {
	
	private static String SPREADSHEET = "@spreadsheet";
	
	private SheetCtrlWidget sheetCtrl;
	public SpreadsheetWidget() {
		this(JQuery.$(SPREADSHEET));
	}
	
	public SpreadsheetWidget(JQuery jquery) {
		super(jquery);
		sheetCtrl = new SheetCtrlWidget(this);
	}
	
	public SheetCtrlWidget getSheetCtrl() {
		return sheetCtrl;
	}
}
