/* Editor.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 15, 2010 7:09:49 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.cell;

import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * A spreadsheet editor helper for onCopy, onPaste, onCut event.
 * @author Sam
 *
 */
public final class EditHelper {
	private EditHelper(){}
	
	private final static String KEY_IS_CUT = "org.zkoss.zss.app.cell.editHelper.isCut";
	private final static String KEY_SRC_SHEET = "org.zkoss.zss.app.cell.editHelper.sourceSheet";
	private final static String KEY_SRC_RANGE = "org.zkoss.zss.app.cell.editHelper.SourceRange";

	/*
	public static void doCut(Spreadsheet ss) {
		if (ss.getSelection() == null)
			return;
		
		ss.setAttribute(KEY_IS_CUT, Boolean.valueOf(true));
		setSource(ss);
	}
	
	public static void doCopy(Spreadsheet ss) {
		if (ss.getSelection() == null)
			return;
		
		ss.setAttribute(KEY_IS_CUT, Boolean.valueOf(false));
		setSource(ss);
	}
	
	public static void clearCutOrCopy(Spreadsheet ss) {
		clearSource(ss);
	}
	*/
	
	/**
	 * Returns whether to cut source range or not
	 * <p> Default: false
	 * @return
	 */
	private static boolean isCut(Spreadsheet ss) {
		return Boolean.valueOf( (Boolean)ss.getAttribute(KEY_IS_CUT) ); //if attr is null return false
	}
	
	/**
	 * Sets source sheet, source range, highlight range
	 * @param ss
	 */
	/*
	private static void setSource(Spreadsheet ss) {
		ss.setAttribute(KEY_SRC_SHEET, ss.getSelectedSheet());
		ss.setAttribute(KEY_SRC_RANGE, ss.getSelection());
		
		Rect sel = ss.getSelection();
		if (sel.getBottom() >= ss.getMaxrows())
			sel.setBottom(ss.getMaxrows() - 1);
		if (sel.getRight() >= ss.getMaxcolumns())
			sel.setRight(ss.getMaxcolumns() - 1);
		ss.setHighlight(sel);
		ss.smartUpdate("copysrc", true);
	}
	*/
	
	/**
	 * Clear source sheet, source range, highlight range
	 * @param ss
	 */
	/*
	private static void clearSource(Spreadsheet ss) {
		ss.setAttribute(KEY_SRC_SHEET, null);
		ss.setAttribute(KEY_SRC_RANGE, null);		
		ss.setHighlight(null);
		ss.smartUpdate("copysrc", false);
	}
	*/
	
	/**
	 * Returns the source sheet to copy.
	 * @param ss
	 * @return sheet
	 */
	/*
	public static Worksheet getSourceSheet(Spreadsheet ss) {
		return (Worksheet)ss.getAttribute(KEY_SRC_SHEET);
	}
	*/
	
	/**
	 * Returns the source range to copy.
	 * @return rect
	 */
	/*
	public static Rect getSourceRange(Spreadsheet ss) {
		return (Rect)ss.getAttribute(KEY_SRC_RANGE);
	}
	*/
	
	/**
	 * Execute paste function use default setting
	 * @param ss
	 * @param pasteType
	 * @param pasteOp
	 * @param skipBlanks
	 * @param transpose
	 */
	/*
	public static void doPaste(Spreadsheet ss) {
		Worksheet srcSheet = getSourceSheet(ss);
		Rect dstRange = ss.getSelection();
		if (srcSheet != null) {
			final Range rng = Utils.pasteSpecial(getSourceSheet(ss),
					getSourceRange(ss), 
					ss.getSelectedSheet(), 
					dstRange.getTop(),
					dstRange.getLeft(),
					dstRange.getBottom(),
					dstRange.getRight(),
					getDefaultPasteType(), 
					getDefaultPasteOperation(), 
					false, false);
			if (rng == null) { //cannot paste
				ss.setSelection(dstRange);
				ss.focus();
				return;
			}
			clearHighlightIfNeed(ss);
			clearCutRangeIfNeed(ss);
			
			int maxColumn = ss.getMaxcolumns();
			int lastColumn = rng.getLastColumn();
			int maxRow = ss.getMaxrows();
			int lastRow = rng.getLastRow();

			ss.setSelection(new Rect(rng.getColumn(), rng.getRow(), 
					lastColumn < maxColumn ? lastColumn : maxColumn - 1, 
					lastRow < maxRow ? lastRow : maxRow - 1));
			ss.focus();
		}
		//TODO : test if needed
		//remSheet.unmergeCells(srcLeft, srcTop, srcRight, srcBottom);
	}
	*/
	
	/*
	private static void clearHighlightIfNeed(Spreadsheet ss) {
		if (isCut(ss))
			ss.setHighlight(null);
	}
	
	private static void clearCutRangeIfNeed(Spreadsheet ss) {
		if (!isCut(ss))
			return;

		int maxCol = ss.getMaxcolumns();
		int maxRow = ss.getMaxrows();
		Worksheet srcSheet = getSourceSheet(ss);
		Worksheet dstSheet = ss.getSelectedSheet();
		Rect srcRange = getSourceRange(ss);
		int srcLeft = srcRange.getLeft();
		int srcRight = srcRange.getRight();
		srcRight = srcRight < maxCol ? srcRight : maxCol;
		
		int srcTop = srcRange.getTop();
		int srcBottom = srcRange.getBottom();
		srcBottom = srcBottom < maxRow ? srcBottom : maxRow;
		
		int dstLeft = ss.getSelection().getLeft();
		int dstTop = ss.getSelection().getTop();
		int dstRight = dstLeft + (srcRight - srcLeft);
		dstRight = dstRight < maxCol ? dstRight : maxCol;
		
		int dstBottom = dstTop + (srcBottom - srcTop);
		dstBottom = dstBottom < maxRow ? dstBottom : maxRow;
		boolean sameSheet = ss.indexOfSheet(srcSheet) == ss.indexOfSheet(dstSheet);
		final CellStyle defaultStyle = ss.getBook().createCellStyle();
		for (int row = srcTop; row <= srcBottom; row++) {
			for (int col = srcLeft; col <= srcRight; col++) {
				if( sameSheet && (row >= dstTop && row <= dstBottom && col >= dstLeft && col <= dstRight) )
					continue;
				Range rng = Ranges.range(srcSheet, row, col);
				rng.setEditText(null);
				rng.setStyle(defaultStyle);
			}
		}
		
		ss.setAttribute(KEY_SRC_SHEET, null);
		ss.setAttribute(KEY_SRC_RANGE, null);
	}
	
	public static void onPasteSpecial(Spreadsheet ss, Worksheet srcSheet, Rect srcRect, int pasteType, int pasteOperation, boolean skipBlanks, boolean transpose){
		if (srcSheet != null) {
			final Rect dst = ss.getSelection();
			final Range rng = Utils.pasteSpecial(srcSheet, 
					srcRect, 
					ss.getSelectedSheet(), 
					dst.getTop(),
					dst.getLeft(),
					dst.getBottom(),
					dst.getRight(),
					pasteType, 
					pasteOperation, 
					skipBlanks, transpose);
			if (rng == null) { //cannot paste
				ss.setSelection(dst);
				ss.focus();
				return;
			}
			clearHighlightIfNeed(ss);
			clearCutRangeIfNeed(ss);
			ss.setSelection(new Rect(rng.getColumn(), rng.getRow(), rng.getLastColumn(), rng.getLastRow()));
			ss.focus();
		}
	}
	*/
	
	/**
	 * Execute paste function base on event's parameter
	 * <p> Parameter can indicate on either paste type or paste operation or transpose
	 * <p> If parameter indicate pasteSpecial, will open a paste special window dialog
	 * <p> If parameter can't find a match, will use default value. 
	 * @param event
	 */
	/*
	public static void onPasteEventHandler(Spreadsheet spreadsheet, String operation) {
		if (spreadsheet == null || getSourceRange(spreadsheet) == null || operation == null) {
//			try {
//				Messagebox.show("Please select a range");
//			} catch (InterruptedException e) {
//			}
			return;
		}

		onPasteSpecial(spreadsheet, getPasteType(operation), getPasteOperation(operation), false, isTranspose(operation));
	}
	*/
	
	public static PasteType getDefaultPasteType() {
		return Range.PasteType.PASTE_ALL;
	}
	
	/**
	 * Returns the paste type base on i3 label (I18N)
	 * <p> Default: returns {@link #Range.PASTE_ALL}, if no match
	 * @return
	 */
	/*
	public static int getPasteType(String type) {
		if (type == null 
				|| "paste".equals(type)
				|| "all".equals(type) )
			return Range.PASTE_ALL;
		
		if ( "allExcpetBorder".equals(type) ) {
			return Range.PASTE_ALL_EXCEPT_BORDERS;
		} else if ( "columnWidth".equals(type) ) {
			return Range.PASTE_COLUMN_WIDTHS;
		} else if ( "comment".equals(type) ) {
			return Range.PASTE_COMMENTS;
		} else if ( "formula".equals(type) ) {
			return Range.PASTE_FORMULAS;
		} else if ( "formulaWithNumFmt".equals(type) ) {
			return Range.PASTE_FORMULAS_AND_NUMBER_FORMATS;
		} else if ( "value".equals(type) ) {
			return Range.PASTE_VALUES;
		} else if ( "valueWithNumFmt".equals(type) ) {
			return Range.PASTE_VALUES_AND_NUMBER_FORMATS;
		} else if ( "format".equals(type) ) {
			return Range.PASTE_FORMATS;
		} else if ( "validation".equals(type) ) {
			return Range.PASTE_VALIDATAION;
		}
		
		return Range.PASTE_ALL;
	}
	
	
	public static int getDefaultPasteOperation() {
		return Range.PASTEOP_NONE;
	}
	*/
	
	/**
	 * Returns the paste operation base on i3-label, if no match
	 * <p> Default: returns {@link #Range.PASTEOP_NONE}, if no match 
	 * @param operation
	 * @return
	 */
	/*
	public static int getPasteOperation(String operation) {
		if (operation == null || "none".equals(operation) )
			return Range.PASTEOP_NONE;
		if ( "add".equals(operation) ) {
			return Range.PASTEOP_ADD;
		} else if ( "sub".equals(operation) ) {
			return Range.PASTEOP_SUB;
		} else if ( "mul".equals(operation) ) {
			return Range.PASTEOP_MUL;
		} else if ( "divide".equals(operation) ) {
			return Range.PASTEOP_DIV;
		}
		return Range.PASTEOP_NONE;
	}
	*/
	
	/**
	 * Returns whether transpose or not
	 * <p> Default: returns false
	 * @param trans
	 * @return 
	 */
	/*
	public static boolean isTranspose(String trans) {
		if (trans == null || !"transpose".equals(trans))
			return false;
		return true;
	}
	*/
	
	/*
	public static void createpPasteSpecialDialog(Spreadsheet spreadsheet, Component parent) {
		Executions.createComponents(
				Consts._PasteSpecialDialog_zul, 
				parent, 
				Zssapps.newSpreadsheetArg(spreadsheet));
	}
	*/
}