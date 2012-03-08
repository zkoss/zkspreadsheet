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

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;

/**
 * Internal Use Only.
 * @author Sam
 *
 */
public class CellVisitorContext {

	private int row;
	private int col;
	private Worksheet sheet;
	private Book book;
	/**
	 * @param sheet
	 * @param row
	 * @param col
	 */
	public CellVisitorContext(Worksheet sheet, int row, int col) {
		this.sheet = sheet;
		this.row = row;
		this.col = col;
		book  = (Book) sheet.getWorkbook();
	}	
	
	/**
	 * Returns the {@link Font} of the current cell
	 * @return  the {@link Font} of the current cell
	 */
	public Font getFont() {
		CellStyle cs = getOrCreateCell().getCellStyle();
		return book.getFontAt((short)cs.getFontIndex());
	}
	
	/**
	 * Returns the {@link CellStyle} of the current cell
	 * @return the {@link CellStyle} of the current cell
	 */
	public CellStyle getCellStyle() {
		return getOrCreateCell().getCellStyle();
	}
	
	/**
	 * Copy the current cell's style setting, returns a new {@link CellStyle}
	 * @return CellStyle
	 */
	public CellStyle cloneCellStyle() {
		CellStyle newCellStyle = book.createCellStyle();
		newCellStyle.cloneStyleFrom(getOrCreateCell().getCellStyle());
		return newCellStyle;
	}
	
	public Font getOrCreateFont(short boldWeight, Color color, short fontHeight, java.lang.String name, 
			boolean italic, boolean strikeout, short typeOffset, byte underline) {
		return BookHelper.getOrCreateFont(book, 
				boldWeight, color, fontHeight, name, 
				italic, strikeout, typeOffset, underline);
	}

	/**
	 * Returns whether font is italic or not.
	 * @return whether font is italic or not.
	 */
	public boolean isItalic() {
		return getFont().getItalic();
	}
	
	/**
	 * Returns whether font is bold or not.
	 * @return whether font is bold or not.
	 */
	public boolean isBold() {
		return getFont().getBoldweight() == Font.BOLDWEIGHT_BOLD;
	}

	/**
	 * Returns font height
	 * @return font height
	 */
	public short getFontHeight() {
		return getFont().getFontHeight();
	}
	
	/**
	 * Returns font family
	 * @return font family
	 */
	public String getFontFamily() {
		return getFont().getFontName();
	}
	
	public String getFontColor() {
		String color = BookHelper.getFontHTMLColor(getOrCreateCell(), getFont());
		return color == null || BookHelper.AUTO_COLOR.equals(color) ? "#000000" : color;
	}

	/**
	 * Returns the cell's alignment
	 * @return the cell's alignment
	 */
	public short getAlignment() {
		return getOrCreateCell().getCellStyle().getAlignment();
	}
	
	/**
	 * Returns the cell's vertical alignment
	 * @return the cell's alignment
	 */
	public short getVerticalAlignment() {
		return getOrCreateCell().getCellStyle().getVerticalAlignment();
	}
	
	/**
	 * Returns the range
	 * @return the range
	 */
	public Range getRange() {
		return Utils.getRange(sheet, row, col);
	}

	/**
	 * Returns cell, if there's no cell, will create cell and return
	 * @return cell of this context
	 */
	public Cell getOrCreateCell() {
		return Utils.getOrCreateCell(sheet, row, col);
	}
	
	/**
	 * Returns cell if there has cell
	 * @return cell if exists; otherwise null
	 */
	public Cell getCell() {
		return Utils.getCell(sheet, row, col);
	}
	
	public Worksheet getSheet() {
		return sheet;
	}
	
	public Book getBook() {
		return book;
	}
	
	public int getRowIndex() {
		return row;
	}
	
	public int getColumnIndex() {
		return col;
	}
	
	public boolean isWrapText() {
		return getCellStyle().getWrapText();
	}
	
	public short getFormatIndex() {
		return getCellStyle().getDataFormat();
	}
	
	public boolean getLocked() {
		return getCellStyle().getLocked();
	}
}