/* CellVisitorContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 17, 2010 2:47:16 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;

/**
 * @author Sam
 *
 */
public class CellVisitorContext {

	int row;
	int col;
	Sheet sheet;
	Cell cell;
	Book book;
	/**
	 * @param sheet
	 * @param row
	 * @param col
	 */
	public CellVisitorContext(Sheet sheet, int row, int col) {
		this.sheet = sheet;
		this.row = row;
		this.col = col;
		cell = Utils.getOrCreateCell(sheet, row, col);
		book  = (Book) sheet.getWorkbook();
	}	
	
	/**
	 * Returns the {#link Font} of the current cell
	 * @return
	 */
	public Font getFont() {
		CellStyle cs = cell.getCellStyle();
		return book.getFontAt((short)cs.getFontIndex());
	}
	
	/**
	 * Returns the {#link CellStyle} of the current cell
	 * @return
	 */
	public CellStyle getCellStyle() {
		return cell.getCellStyle();
	}
	
	/**
	 * Copy the current cell's style setting, returns a new {#link CellStyle}
	 * @return CellStyle
	 */
	public CellStyle cloneCellStyle() {
		CellStyle newCellStyle = book.createCellStyle();
		newCellStyle.cloneStyleFrom(cell.getCellStyle());
		return newCellStyle;
	}
	
	public Font getOrCreateFont(short boldWeight, short color, short fontHeight, java.lang.String name, 
			boolean italic, boolean strikeout, short typeOffset, byte underline) {
		return BookHelper.getOrCreateFont(book, 
				boldWeight, color, fontHeight, name, 
				italic, strikeout, typeOffset, underline);
	}

	/**
	 * Returns whether font is italic or not.
	 * @return
	 */
	public boolean isItalic() {
		return getFont().getItalic();
	}
	
	/**
	 * Returns whether font is bold or not.
	 * @return
	 */
	public boolean isBold() {
		return getFont().getBoldweight() == Font.BOLDWEIGHT_BOLD;
	}

	/**
	 * Returns font height
	 * @return
	 */
	public short getFontHeight() {
		return getFont().getFontHeight();
	}
	
	/**
	 * Returns font family
	 * @return
	 */
	public String getFontFamily() {
		return getFont().getFontName();
	}
	
	/**
	 * Returns the cell's alignment
	 * @return
	 */
	public short getAlignment() {
		return cell.getCellStyle().getAlignment();
	}
	
	/**
	 * Returns the range
	 * @return
	 */
	public Range getRange() {
		return Utils.getRange(sheet, row, col);
	}

	/**
	 * @return
	 */
	public Cell getCell() {
		return cell;
	}
}