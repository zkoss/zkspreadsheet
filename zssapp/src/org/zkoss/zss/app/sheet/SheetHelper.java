/* SheetHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 7:05:07 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.sheet;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Messagebox;

/**
 * @author Sam
 *
 */
public final class SheetHelper {
	private SheetHelper(){};
	
	/**
	 * Returns the shift sheet's index
	 * @param spreadsheet
	 * @return -1 if fail to shift sheet
	 */
	public static int shiftSheetLeft(Spreadsheet spreadsheet) {
		//TODO: it should shift all, not just itself
		Book book = spreadsheet.getBook();
		String name = spreadsheet.getSelectedSheet().getSheetName();
		Sheet sheet = spreadsheet.getSelectedSheet();
		int index = book.getSheetIndex(sheet);
		if (index > 0) {
			int newIdx = index - 1;
			book.setSheetOrder(name, index - 1);
			return newIdx;
		}
		return -1;
	}
	
	/**
	 * Returns the shift sheet's index
	 * @param spreadsheet
	 * @return -1 if fail to shift sheet
	 */
	public static int shiftSheetRight(Spreadsheet spreadsheet) {
		//TODO: it should shift all, not just itself
		Book book = spreadsheet.getBook();
		String name = spreadsheet.getSelectedSheet().getSheetName();
		Sheet sheet = spreadsheet.getSelectedSheet();
		int index = book.getSheetIndex(sheet);
		if (index < book.getNumberOfSheets() - 1) {
			int newIdx = index + 1;
			book.setSheetOrder(name, newIdx);
			return newIdx;
		}
		return -1;
	}
	
	/**
	 * 
	 * @param spreadsheet
	 * @return new index, -1 if delete all sheet
	 */
	public static int deleteSheet(Spreadsheet spreadsheet) {
		//TODO: it should shift all, not just itself
		Book book = spreadsheet.getBook();
		int index = book.getSheetIndex(spreadsheet.getSelectedSheet());
		book.removeSheetAt(index);
		int sheetCount = book.getNumberOfSheets();
		
		//TODO: can remove all sheets ?
		if (index < sheetCount)
			return index;
		return -1;
	}
	
	/**
	 * Rename sheet
	 * @param spreadsheet
	 * @param name
	 * @return
	 */
	public static int renameSheet(Spreadsheet spreadsheet, String name) {
		Book book = spreadsheet.getBook();
		final Sheet selsheet  = spreadsheet.getSelectedSheet();
		Sheet sheet = book.getSheet(name);
		if(sheet != null)
			return -1;
		final int index = book.getSheetIndex(selsheet);
		book.setSheetName(index, name);
		return index;
	}
	
	public static Rect getSpreadsheetMaxSelection(Spreadsheet spreadsheet) {
		Rect selection = spreadsheet.getSelection();
		if (selection.getBottom() >= spreadsheet.getMaxrows())
			selection.setBottom(spreadsheet.getMaxrows() - 1);
		if (selection.getRight() >= spreadsheet.getMaxcolumns())
			selection.setRight(spreadsheet.getMaxcolumns() - 1);
		return selection;
	}
}
