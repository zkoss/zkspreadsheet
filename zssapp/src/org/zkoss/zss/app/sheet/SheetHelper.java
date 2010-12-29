/* SheetHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 7:05:07 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.sheet;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;

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
		final Book book = spreadsheet.getBook();
		if (book != null) {
			String name = spreadsheet.getSelectedSheet().getSheetName();
			Sheet sheet = spreadsheet.getSelectedSheet();
			int index = book.getSheetIndex(sheet);
			if (index > 0) {
				int newIdx = index - 1;
				book.setSheetOrder(name, index - 1);
				return newIdx;
			}
		}
		return -1;
	}
	
	/**
	 * Returns the shift sheet's index
	 * @param spreadsheet
	 * @return -1 if fail to shift sheet
	 */
	public static int shiftSheetRight(Spreadsheet spreadsheet) {
		final Book book = spreadsheet.getBook();
		if (book != null) {
			String name = spreadsheet.getSelectedSheet().getSheetName();
			Sheet sheet = spreadsheet.getSelectedSheet();
			int index = book.getSheetIndex(sheet);
			if (index < book.getNumberOfSheets() - 1) {
				int newIdx = index + 1;
				book.setSheetOrder(name, newIdx);
				SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
				ctrl.getWidgetHandler().invaliate();
				return newIdx;
			}
		}
		return -1;
	}
	
	/**
	 * 
	 * @param spreadsheet
	 * @return new index, -1 if delete sheet fail
	 */
	public static int deleteSheet(Spreadsheet spreadsheet) {
		final Book book = spreadsheet.getBook();
		if (book != null) {
			//Note. Sheet must contain at least one sheet
			int sheetCount = book.getNumberOfSheets();
			if (sheetCount == 1)
				return -1;
			
			int index = book.getSheetIndex(spreadsheet.getSelectedSheet());
			book.removeSheetAt(index);
			sheetCount = book.getNumberOfSheets();
			
			if (index < sheetCount) {
				//shift right
				return index;
			} else {
				//shift left
				return index - 1;
			}
		}
		return -1;
	}
	
	/**
	 * Rename sheet
	 * @param spreadsheet
	 * @param name
	 * @return
	 */
	public static int renameSheet(Spreadsheet spreadsheet, String name) {
		final Book book = spreadsheet.getBook();
		if (book != null) {
			final Sheet selsheet  = spreadsheet.getSelectedSheet();
			Sheet sheet = book.getSheet(name);
			if(sheet != null)
				return -1;
			final int index = book.getSheetIndex(selsheet);
			book.setSheetName(index, name);
			BookHelper.getOrCreateRefBook(book).setSheetName(selsheet.getSheetName(), name);
			return index;
		}
		return -1;
	}
	
	public static Rect getSpreadsheetMaxSelection(Spreadsheet spreadsheet) {
		Rect selection = spreadsheet.getSelection();// selection is cloned
		if (selection.getBottom() >= spreadsheet.getMaxrows())
			selection.setBottom(spreadsheet.getMaxrows() - 1);
		if (selection.getRight() >= spreadsheet.getMaxcolumns())
			selection.setRight(spreadsheet.getMaxcolumns() - 1);
		return selection;
	}
}
