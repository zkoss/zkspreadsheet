/* CellHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 15, 2010 7:09:49 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.cell;

//import org.zkoss.poi.ss.usermodel.Cell;
//import org.zkoss.poi.ss.usermodel.CellStyle;
//import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
//import org.zkoss.zss.ui.impl.Utils;

public final class CellHelper {
//	private CellHelper(){};
//	
//	/**
//	 * Returns cell font family
//	 * @param cell
//	 * @return
//	 */
//	public static Font getFont(Cell cell) {
//		if (cell == null)
//			return null;
//		
//		XBook book = (XBook)cell.getSheet().getWorkbook();
//		CellStyle cellStyle = cell.getCellStyle();
//		return book.getFontAt(cellStyle.getFontIndex());
//	}
//	
//	public static boolean isBold(Font font) {
//		return font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
//	}
//	
//	public static boolean isAlignLeft(Cell cell) {
//		CellStyle cellStyle = cell.getCellStyle();
//		return cellStyle.getAlignment() == CellStyle.ALIGN_LEFT;
//	}
//	
//	public static boolean isAlignCenter(Cell cell) {
//		CellStyle cellStyle = cell.getCellStyle();
//		return cellStyle.getAlignment() == CellStyle.ALIGN_CENTER;
//	}
//
//	public static boolean isAlignRight(Cell cell) {
//		CellStyle cellStyle = cell.getCellStyle();
//		return cellStyle.getAlignment() == CellStyle.ALIGN_RIGHT;
//	}
//	
//	public static String getFontHTMLColor(Cell cell, Font font) {
//		XBook book = (XBook)cell.getSheet().getWorkbook();
////		String color = BookHelper.getFontHTMLColor(book, font);
//		String color = BookHelper.getFontHTMLColor(cell, font);
//		if (color == null || BookHelper.AUTO_COLOR.equals(color))
//			return "#000000";
//		return  color;
//	}
//	
//	public static String getBackgroundHTMLColor(Cell cell) {
//		CellStyle cellStyle = cell.getCellStyle();
//		XBook book = (XBook)cell.getSheet().getWorkbook();
//		//bug#ZSS-34: cell background color does not show in excel
//		String color = cellStyle.getFillPattern() != CellStyle.NO_FILL ? 
//				BookHelper.colorToHTML(book, cellStyle.getFillForegroundColorColor()) : null;
//		if (color == null || BookHelper.AUTO_COLOR.equals(color))
//			return "#FFFFFF";
//		return color;
//	}
//
//	public static void clearContent(Spreadsheet spreadsheet, Rect rect) {
//		Utils.visitCells(spreadsheet.getSelectedSheet(), 
//				rect, 
//				new CellVisitor() {
//					@Override
//					public void handle(CellVisitorContext context) {
//						Cell cell = context.getCell();
//						if (cell != null) {
//							context.getRange().setEditText(null);
//						}
//					}
//		});
//	}
//	
//	public static void clearStyle(Spreadsheet spreadsheet, Rect rect) {
//		XRanges.range(
//				spreadsheet.getSelectedSheet(), 
//				rect.getTop(), rect.getLeft(),rect.getBottom(), rect.getRight()).
//					setStyle(spreadsheet.getXBook().createCellStyle());
//	}
//	
//	/**
//	 * Delete current selection area and shift cells below the selection area up
//	 * @param XSheet the current sheet
//	 * @param rect the selection rectangle
//	 */
//	public static void shiftCellUp(XSheet sheet, Rect rect) {
//		final XRange rng = XRanges.range(sheet, rect.getTop(), rect.getLeft(), rect.getBottom(), rect.getRight());
//		rng.delete(XRange.SHIFT_UP);
//	}
//	
//	/**
//	 * Insert a new cell and shift cells right
//	 * @param XSheet the current sheet
//	 * @param rect the selection rectangle
//	 */
//	public static void shiftCellRight(XSheet sheet, Rect rect) {
//		final XRange rng = XRanges.range(sheet, rect.getTop(), rect.getLeft(), rect.getBottom(), rect.getRight());
//		rng.insert(XRange.SHIFT_RIGHT, XRange.FORMAT_RIGHTBELOW);
//	}
//	
//	/**
//	 * Insert a new cell and shift original cells down
//	 * @param XSheet the current sheet
//	 * @param rect the selection rectangle
//	 */
//	public static void shiftCellDown(XSheet sheet, Rect rect) {
//		final XRange rng = XRanges.range(sheet, rect.getTop(), rect.getLeft(), rect.getBottom(), rect.getRight());
//		rng.insert(XRange.SHIFT_DOWN, XRange.FORMAT_LEFTABOVE);
//	}
//	
//	/**
//	 * Delete current cell and shift cells up beside it.
//	 * @param XSheet the current sheet
//	 * @param rect the selection rectangle
//	 */
//	public static void shiftCellLeft(XSheet sheet, Rect rect) {
//		final XRange rng = XRanges.range(sheet, rect.getTop(), rect.getLeft(), rect.getBottom(), rect.getRight());
//		rng.delete(XRange.SHIFT_LEFT);
//	}
//	
//	public static void shiftEntireRowDown(XSheet sheet, int top, int bottom) {
//		XRanges.range(sheet, top, 0, bottom, 0).getRows().insert(XRange.SHIFT_DOWN, XRange.FORMAT_LEFTABOVE);
//	}
//	
//	public static void shiftEntireRowUp(XSheet sheet, int top, int bottom) {
//		XRanges.range(sheet, top, 0, bottom, 0).getRows().delete(XRange.SHIFT_UP);
//	}
//
//	public static void shiftEntireColumnRight(XSheet sheet, int left, int right) {
//		XRanges.range(sheet, 0, left, 0, right).getColumns().insert(XRange.SHIFT_RIGHT, XRange.FORMAT_RIGHTBELOW);
//	}
//	
//	public static void shiftEntireColumnLeft(XSheet sheet, int left, int right) {
//		XRanges.range(sheet, 0, left, 0, right).getColumns().delete(XRange.SHIFT_LEFT);
//	}
//	
//	
//	public static void sortAscending(XSheet sheet, Rect rect) {
//		Utils.sort(sheet, rect,
//				null, null, null, false, false, false);
//	}
//	
//	public static void sortDescending(XSheet sheet, Rect rect) {
//		Utils.sort(sheet, rect,
//				null, new boolean[] { true }, null, false, false, false);
//	}
//
//	public static void autoFilter(XSheet sheet,
//			Rect rect) {
//		Utils.autoFilter(sheet,rect);
//	}
}