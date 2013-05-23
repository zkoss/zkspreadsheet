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

//import org.zkoss.poi.ss.usermodel.BorderStyle;
//import org.zkoss.poi.ss.usermodel.Cell;
//import org.zkoss.poi.ss.usermodel.CellStyle;
//import org.zkoss.poi.ss.usermodel.Font;
//import org.zkoss.poi.ss.usermodel.FontUnderline;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.CellOperationUtil.CellStyleApplier;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.UnitUtil;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.sheet.SheetHelper;
//import org.zkoss.zss.model.sys.XRanges;
//import org.zkoss.zss.model.sys.XSheet;
//import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
//import org.zkoss.zss.ui.impl.CellVisitor;
//import org.zkoss.zss.ui.impl.CellVisitorContext;
//import org.zkoss.zss.ui.impl.Utils;

/**
 * @author Ian Tsai / Sam
 *
 */
public class SSRectCellStyle implements org.zkoss.zss.app.zul.ctrl.CellStyleApplier {

	Spreadsheet spreadsheet;
	private Font font;
	private org.zkoss.zss.api.model.CellStyle style;
//	private Cell cell;
	int row,col;
	
	/**
	 * @param cell
	 * @param spreadsheet
	 */
	public SSRectCellStyle(Spreadsheet spreadsheet, int row, int col) {
		super();
		this.row = row;
		this.col = col;
		this.spreadsheet = spreadsheet;

		resetFont();
	}
	
	private void resetFont(){
		style = Ranges.range(spreadsheet.getSelectedSheet(),row,col).getCellStyle();
		font = style.getFont();
	}
	
	public void setFontSize(int size) {
		Sheet sheet = spreadsheet.getSelectedSheet();
		Rect rect = SheetHelper.getVisibleSelection(spreadsheet);

		Range r = Ranges.range(sheet,rect);
		CellOperationUtil.applyFontSize(r, (short)size);
		
		
//		Utils.setFontHeight(sheet, 
//				rect, 
//				getFontHeight(size));	
//		setProperRowHeightByFontSize(sheet, rect, size);
		resetFont();
	}
//	private short getFontHeight(int size) {
//		return (short)(size * 20);
//	}
	
//	private void setProperRowHeightByFontSize(Sheet sheet, Rect rect, int size) {	
//		int tRow = rect.getTop();
//		int bRow = rect.getBottom();
//		int col = rect.getLeft();
//		
//		for (int i = tRow; i <= bRow; i++) {
//			//Note. add extra padding height: 4
//			if ((size + 4) > (Utils.pxToPoint(Utils.twipToPx(BookHelper.getRowHeight(sheet, i))))) {
//				Ranges.range(sheet, i, col).setRowHeight(size + 4);
//			}
//		}
//	}
	
	public void setFontFamily(String family) {
		//TODO: use Utils.setFontFamily will fire many SSDataEvent 
		
		Range r = Ranges.range( spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyFontName(r, family);
//		Utils.setFontFamily(spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet),
//				family);
		resetFont();
	}
	
	public void setBold(boolean bold) {
		//TODO: use Utils.setFontBold will fire many SSDataEvent 
		
		Range r = Ranges.range( spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyFontBoldweight(r, bold?Boldweight.BOLD:Boldweight.NORMAL);
//		Utils.setFontBold(spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet),
//				bold);
		resetFont();
	}
	
	public String getFontFamily() {
		return font.getFontName();
	}
	
	public boolean isBold() {
		Boldweight bw = font.getBoldweight();
		resetFont();
		return  bw == Boldweight.NORMAL;
	}
	
	public int getFontSize() {
		return UnitUtil.twipToPoint(font.getFontHeight());
	}

	public Alignment getAlignment() {
		return style.getAlignment();
	}

	public String getCellColor() {
		return style.getBackgroundColor().getHtmlColor();
	}

	public String getFontColor() {
		return font.getColor().getHtmlColor();
	}

	public boolean isItalic() {
		return font.isItalic();
	}

	public boolean isStrikethrough() {
		return font.isStrikeout();
	}

	public Underline getUnderline() {
		return font.getUnderline();
	}

	public void setAlignment(Alignment alignment) {
		//TODO: Utils.setAlignment will fire many SSDataEvent
		
		Range r = Ranges.range( spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyAlignment(r, alignment);
		
		
//		Utils.setAlignment(spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet), 
//				(short)alignment);
		resetFont();
	}
	

	public void setVerticalAlignment(final VerticalAlignment alignment) {
		Range r = Ranges.range( spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyVerticalAlignment(r, alignment);
//		Utils.visitCells(spreadsheet.getSelectedSheet(), SheetHelper.getVisibleSelection(spreadsheet), new CellVisitor(){
//			@Override
//			public void handle(CellVisitorContext context) {
//				final short srcAlign = context.getVerticalAlignment();
//
//				if (srcAlign != alignment) {
//					CellStyle newStyle = context.cloneCellStyle();
//					newStyle.setVerticalAlignment((short)alignment);
//					context.getRange().setStyle(newStyle);
//				}
//			}});
		resetFont();
	}

	public VerticalAlignment getVerticalAlignment() {
		return style.getVerticalAlignment();
	}

	public void setBorder(ApplyBorderType position, BorderType borderType, String color) {
		//TODO: Utils.setBorder will fire many SSDataEvent
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyBorder(r, position, borderType, color);
//		Utils.setBorder(spreadsheet.getSelectedSheet(), 
//				SheetHelper.getSpreadsheetMaxSelection(spreadsheet), 
//				(short)borderPosition, borderStyle, color);
//		
		resetFont();
	}

	public void setCellColor(String color) {
		//TODO: Utils.setBackgroundColor will fire many SSDataEvent
		
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyBackgroundColor(r, color);
		
//		Utils.setBackgroundColor(
//				spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet), 
//				color);
		resetFont();
	}

	public void setFontColor(String color) {
		//TODO: Utils.setFontColor will fire many SSDataEvent
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyFontColor(r, color);
		
//		Utils.setFontColor(
//				spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet), 
//				color);
		resetFont();
	}

	public void setItalic(boolean italic) {
		//TODO: Utils.setFontItalic will fire many SSDataEvent
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyFontItalic(r,italic);
//		Utils.setFontItalic(
//				spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet),
//				italic);
		resetFont();
	}

	public void setStrikethrough(boolean strikethrough) {
		//TODO: Utils.setFontStrikeout will fire many SSDataEvent
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyFontStrikeout(r, strikethrough);
//		Utils.setFontStrikeout(
//				spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet), 
//				strikethrough);
		resetFont();
	}

	public void setUnderline(Underline underlineStyle) {
		
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyFontUnderline(r, underlineStyle);
//		FontUnderline underline = FontUnderline.NONE;
//		if (underlineStyle == UNDERLINE_SINGLE)
//			underline = FontUnderline.SINGLE;
//		//TODO: Utils.setFontUnderline will fire many SSDataEvent
//		Utils.setFontUnderline(
//				spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet),
//				underline.getByteValue());
		resetFont();
	}
	
	public boolean isWrapText() {
		return style.isWrapText();
	}

	public void setWrapText(boolean wrapped) {
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyWrapText(r, wrapped);
//		Utils.setWrapText(spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet), wrapped);
		resetFont();
	}
	
//	public String toString(){
//		return spreadsheet.getColumntitle(column)+ 
//		cell.getRowIndex();
//	}

	@Override
	public void setLocked(final boolean locked) {
		Range r = Ranges.range(spreadsheet.getSelectedSheet(),SheetHelper.getVisibleSelection(spreadsheet));
		CellOperationUtil.applyCellStyle(r, new CellStyleApplier() {
			public boolean ignore(Range cellRange,CellStyle oldCellstyle) {
				return oldCellstyle.isLocked()==locked;
			}

			public void apply(Range cellRange,
					org.zkoss.zss.api.model.CellStyle newCellstyle) {
				newCellstyle.setLocked(locked);
			}
		});
		
//		Utils.setLocked(
//				spreadsheet.getSelectedSheet(), 
//				SheetHelper.getVisibleSelection(spreadsheet), locked);
	}

	@Override
	public boolean getLocked() {
		return style.isLocked();
	}
}