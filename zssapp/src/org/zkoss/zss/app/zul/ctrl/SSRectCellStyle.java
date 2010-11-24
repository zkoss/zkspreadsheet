/* SSRectFontStyle.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 11:35:30 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.FontUnderline;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

/**
 * @author Ian Tsai / Sam
 *
 */
public class SSRectCellStyle implements CellStyle {

	Spreadsheet spreadsheet;
	private Font font;
	private Cell cell;
	
	/**
	 * @param cell
	 * @param spreadsheet
	 */
	public SSRectCellStyle(Cell cell, Spreadsheet spreadsheet) {
		super();
		this.cell = cell;
		this.spreadsheet = spreadsheet;

		resetFont();
	}
	
	private void resetFont(){
		short idx = cell.getCellStyle().getFontIndex();
		font = spreadsheet.getBook().getFontAt(idx);
	}
	
	@Override
	public void setFontSize(int size) {
		
		Sheet sheet = spreadsheet.getSelectedSheet();
		Rect rect = spreadsheet.getSelection();

		Utils.setFontHeight(sheet, 
				rect, 
				getFontHeight(size));	
		setProperRowHeightByFontSize(sheet, rect, size);
		resetFont();
	}
	private short getFontHeight(int size) {
		return (short)(size * 20);
	}
	private void setProperRowHeightByFontSize(Sheet sheet, Rect rect, int size) {	
		int seldFontHeight = getFontHeight(size);
		int tRow = rect.getTop();
		int bRow = rect.getBottom();
		int col = rect.getLeft();
		
		for (int i = tRow; i <= bRow; i++) {
			//Note. the book helper return measured in twips (1/20 of  a point)
			if (seldFontHeight > BookHelper.getRowHeight(sheet, i)) {
				Ranges.range(sheet, i, col).setRowHeight(size + 10);
			}
		}
	}
	
	@Override
	public void setFontFamily(String family) {
		Utils.setFontFamily(spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				family);
		resetFont();
	}
	
	@Override
	public void setBold(boolean bold) {
		Utils.setFontBold(spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				bold);
		resetFont();
	}
	
	@Override
	public String getFontFamily() {
		return font.getFontName();
	}
	
	@Override
	public boolean isBold() {
		short bold = font.getBoldweight();
		return  Font.BOLDWEIGHT_BOLD== bold;
	}
	
	@Override
	public int getFontSize() {
		return font.getFontHeight() / 20;
	}

	@Override
	public int getAlignment() {
		return cell.getCellStyle().getAlignment();
	}

	@Override
	public String getCellColor() {
		return CellHelper.getBackgroundHTMLColor(cell);
	}

	@Override
	public String getFontColor() {
		return CellHelper.getFontHTMLColor(cell, font);
	}

	@Override
	public boolean isItalic() {
		return font.getItalic();
	}

	@Override
	public boolean isStrikethrough() {
		return font.getStrikeout();
	}

	@Override
	public int getUnderline() {
		return (int)font.getUnderline();
	}

	@Override
	public void setAlignment(int alignment) {
		Utils.setAlignment(spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				(short)alignment);
	}

	@Override
	public void setBorder(int borderPosition, BorderStyle borderStyle, String color) {
		Utils.setBorder(spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				(short)borderPosition, borderStyle, color);
		
	}

	@Override
	public void setCellColor(String color) {
		Utils.setBackgroundColor(
				spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				color);
	}

	@Override
	public void setFontColor(String color) {
		Utils.setFontColor(
				spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				color);
	}

	@Override
	public void setItalic(boolean italic) {
		Utils.setFontItalic(
				spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				italic);
	}

	@Override
	public void setStrikethrough(boolean strikethrough) {
		Utils.setFontStrikeout(
				spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				strikethrough);
	}

	@Override
	public void setUnderline(int underlineStyle) {
		
		FontUnderline underline = FontUnderline.NONE;
		if (underlineStyle == UNDERLINE_SINGLE)
			underline = FontUnderline.SINGLE;
		
		Utils.setFontUnderline(
				spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				underline.getByteValue());
	}
	
	public boolean isWrapText() {
		return cell.getCellStyle().getWrapText();
	}

	public void setWrapText(boolean wrapped) {
		Utils.setWrapText(spreadsheet.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), wrapped);
	}
	
	public String toString(){
		return spreadsheet.getColumntitle(cell.getColumnIndex())+ 
		cell.getRowIndex();
	}
}