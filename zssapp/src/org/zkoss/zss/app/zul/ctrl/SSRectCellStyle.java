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
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.FontUnderline;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.CellVisitor;
import org.zkoss.zss.ui.impl.CellVisitorContext;
import org.zkoss.zss.ui.impl.Utils;

/**
 * @author Ian Tsai / Sam
 *
 */
public class SSRectCellStyle implements org.zkoss.zss.app.zul.ctrl.CellStyle {

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
		font = spreadsheet.getXBook().getFontAt(idx);
	}
	
	public void setFontSize(int size) {
		XSheet sheet = spreadsheet.getSelectedXSheet();
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
	
	private void setProperRowHeightByFontSize(XSheet sheet, Rect rect, int size) {	
		int tRow = rect.getTop();
		int bRow = rect.getBottom();
		int col = rect.getLeft();
		
		for (int i = tRow; i <= bRow; i++) {
			//Note. add extra padding height: 4
			if ((size + 4) > (Utils.pxToPoint(Utils.twipToPx(BookHelper.getRowHeight(sheet, i))))) {
				XRanges.range(sheet, i, col).setRowHeight(size + 4);
			}
		}
	}
	
	public void setFontFamily(String family) {
		//TODO: use Utils.setFontFamily will fire many SSDataEvent 
		Utils.setFontFamily(spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				family);
		resetFont();
	}
	
	public void setBold(boolean bold) {
		//TODO: use Utils.setFontBold will fire many SSDataEvent 
		Utils.setFontBold(spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				bold);
		resetFont();
	}
	
	public String getFontFamily() {
		return font.getFontName();
	}
	
	public boolean isBold() {
		short bold = font.getBoldweight();
		resetFont();
		return  Font.BOLDWEIGHT_BOLD == bold;
	}
	
	public int getFontSize() {
		return font.getFontHeight() / 20;
	}

	public int getAlignment() {
		return (int)cell.getCellStyle().getAlignment();
	}

	public String getCellColor() {
		return CellHelper.getBackgroundHTMLColor(cell);
	}

	public String getFontColor() {
		return CellHelper.getFontHTMLColor(cell, font);
	}

	public boolean isItalic() {
		return font.getItalic();
	}

	public boolean isStrikethrough() {
		return font.getStrikeout();
	}

	public int getUnderline() {
		return (int)font.getUnderline();
	}

	public void setAlignment(int alignment) {
		//TODO: Utils.setAlignment will fire many SSDataEvent 
		Utils.setAlignment(spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				(short)alignment);
		resetFont();
	}
	

	public void setVerticalAlignment(final int alignment) {
		Utils.visitCells(spreadsheet.getSelectedXSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet), new CellVisitor(){
			@Override
			public void handle(CellVisitorContext context) {
				final short srcAlign = context.getVerticalAlignment();

				if (srcAlign != alignment) {
					CellStyle newStyle = context.cloneCellStyle();
					newStyle.setVerticalAlignment((short)alignment);
					context.getRange().setStyle(newStyle);
				}
			}});
		resetFont();
	}

	public int getVerticalAlignment() {
		return cell.getCellStyle().getVerticalAlignment();
	}

	public void setBorder(int borderPosition, BorderStyle borderStyle, String color) {
		//TODO: Utils.setBorder will fire many SSDataEvent
		Utils.setBorder(spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				(short)borderPosition, borderStyle, color);
		resetFont();
	}

	public void setCellColor(String color) {
		//TODO: Utils.setBackgroundColor will fire many SSDataEvent
		Utils.setBackgroundColor(
				spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				color);
		resetFont();
	}

	public void setFontColor(String color) {
		//TODO: Utils.setFontColor will fire many SSDataEvent
		Utils.setFontColor(
				spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				color);
		resetFont();
	}

	public void setItalic(boolean italic) {
		//TODO: Utils.setFontItalic will fire many SSDataEvent
		Utils.setFontItalic(
				spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				italic);
		resetFont();
	}

	public void setStrikethrough(boolean strikethrough) {
		//TODO: Utils.setFontStrikeout will fire many SSDataEvent
		Utils.setFontStrikeout(
				spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
				strikethrough);
		resetFont();
	}

	public void setUnderline(int underlineStyle) {
		FontUnderline underline = FontUnderline.NONE;
		if (underlineStyle == UNDERLINE_SINGLE)
			underline = FontUnderline.SINGLE;
		//TODO: Utils.setFontUnderline will fire many SSDataEvent
		Utils.setFontUnderline(
				spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet),
				underline.getByteValue());
		resetFont();
	}
	
	public boolean isWrapText() {
		return cell.getCellStyle().getWrapText();
	}

	public void setWrapText(boolean wrapped) {
		Utils.setWrapText(spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), wrapped);
		resetFont();
	}
	
	public String toString(){
		return spreadsheet.getColumntitle(cell.getColumnIndex())+ 
		cell.getRowIndex();
	}

	@Override
	public void setLocked(boolean locked) {
		Utils.setLocked(
				spreadsheet.getSelectedXSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), locked);
	}

	@Override
	public boolean getLocked() {
		return cell.getCellStyle().getLocked();
	}
}