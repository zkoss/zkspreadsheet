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

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

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
	
	/**
	 * Clear source sheet, source range, highlight range
	 * @param ss
	 */
	private static void clearSource(Spreadsheet ss) {
		ss.setAttribute(KEY_SRC_SHEET, null);
		ss.setAttribute(KEY_SRC_RANGE, null);		
		ss.setHighlight(null);
		ss.smartUpdate("copysrc", false);
	}
	
	/**
	 * Returns the source sheet to copy.
	 * @param ss
	 * @return sheet
	 */
	public static Worksheet getSourceSheet(Spreadsheet ss) {
		return (Worksheet)ss.getAttribute(KEY_SRC_SHEET);
	}
	
	/**
	 * Returns the source range to copy.
	 * @return rect
	 */
	public static Rect getSourceRange(Spreadsheet ss) {
		return (Rect)ss.getAttribute(KEY_SRC_RANGE);
	}
	
	
	//TODO: test copy/cut behavior on excel for overlap cell 
//	private boolean isOverlapMergedCell(Sheet sheet, int top, int left, int bottom, int right) {
//		MergeMatrixHelper mmhelper = ss.getMergeMatrixHelper(sheet);
//		for(final Iterator iter = mmhelper.getRanges().iterator(); iter.hasNext();) {
//			final MergedRect block = (MergedRect) iter.next();
//			int bl = block.getLeft();
//			int br = block.getRight();
//			int bt = block.getTop();
//			int bb = block.getBottom();
//			if (bt <= bottom && bl <= right 
//				&& br >= left && bb >= top)
//				return true;
//		}
//		return false;
//	}
	
	/**
	 * Execute paste function use default setting
	 * @param ss
	 * @param pasteType
	 * @param pasteOp
	 * @param skipBlanks
	 * @param transpose
	 */
	public static void doPaste(Spreadsheet ss) {
		Worksheet srcSheet = getSourceSheet(ss);
		Rect dstRange = ss.getSelection();
		if (srcSheet != null) {
			//TODO: **test overlap merge cell behavior on excel**
//			if(isOverlapMergedCell(sheet, dstTop, dstLeft, dstBottom, dstRight)) { 
//			try {
//				Messagebox.show("cannot change part of merged cell in destination region");
//				return;
//			} catch (InterruptedException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//		}
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
	
	private static void clearHighlightIfNeed(Spreadsheet ss) {
		if (isCut(ss))
			ss.setHighlight(null);
	}
	
	private static void clearCutRangeIfNeed(Spreadsheet ss) {
		if (!isCut(ss))
			return;

		Worksheet srcSheet = getSourceSheet(ss);
		Worksheet dstSheet = ss.getSelectedSheet();
		Rect srcRange = getSourceRange(ss);
		int srcLeft = srcRange.getLeft();
		int srcRight = srcRange.getRight();
		int srcTop = srcRange.getTop();
		int srcBottom = srcRange.getBottom();
		
		int dstLeft = ss.getSelection().getLeft();
		int dstTop = ss.getSelection().getTop();
		int dstRight = dstLeft + (srcRight - srcLeft);
		int maxCol = ss.getMaxcolumns();
		dstRight = dstRight < maxCol ? dstRight : maxCol;
		
		int dstBottom = dstTop + (srcBottom - srcTop);
		int maxRow = ss.getMaxrows();
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
	
	public static void onPasteSpecial(Spreadsheet ss, int pasteType, int pasteOperation, boolean skipBlanks, boolean transpose){
		Worksheet srcSheet = getSourceSheet(ss);
		if (srcSheet != null) {
			final Rect dst = ss.getSelection();
			final Range rng = Utils.pasteSpecial(srcSheet, 
					getSourceRange(ss), 
					ss.getSelectedSheet(), 
					dst.getTop(),
					dst.getLeft(),
					dst.getBottom(),
					dst.getRight(),
					pasteType, 
					pasteOperation, 
					skipBlanks, transpose);
			clearHighlightIfNeed(ss);
			clearCutRangeIfNeed(ss);
			ss.setSelection(new Rect(rng.getColumn(), rng.getRow(), rng.getLastColumn(), rng.getLastRow()));
			ss.focus();
		}
	}
	
	/**
	 * Execute paste function base on event's parameter
	 * <p> Parameter can indicate on either paste type or paste operation or transpose
	 * <p> If parameter indicate pasteSpecial, will open a paste special window dialog
	 * <p> If parameter can't find a match, will use default value. 
	 * @param event
	 */
	public static void onPasteEventHandler(Spreadsheet spreadsheet, String operation) {
		if (spreadsheet == null || getSourceRange(spreadsheet) == null || operation == null) {
//			try {
//				Messagebox.show("Please select a range");
//			} catch (InterruptedException e) {
//			}
			return;
		}

		if (operation.equals(Labels.getLabel("pasteSpecial")) ) {
			Executions.createComponents(Consts._PasteSpecialDialog_zul, null, Zssapps.newSpreadsheetArg(spreadsheet));
		} else {
			onPasteSpecial(spreadsheet, getPasteType(operation), getPasteOperation(operation), false, isTranspose(operation));
		}
	}
	
	public static int getDefaultPasteType() {
		return Range.PASTE_ALL;
	}
	
	/**
	 * Returns the paste type base on i3 label (I18N)
	 * <p> Default: returns {@link #Range.PASTE_ALL}, if no match
	 * @return
	 */
	public static int getPasteType(String i3label) {
		if (i3label == null 
				|| i3label.equals(Labels.getLabel("paste"))
				|| i3label.equals(Labels.getLabel("paste.all")) )
			return Range.PASTE_ALL;
		
		if ( i3label.equals(Labels.getLabel("paste.allExcpetBorder")) ) {
			return Range.PASTE_ALL_EXCEPT_BORDERS;
		} else if ( i3label.equals(Labels.getLabel("paste.columnWidth")) ) {
			return Range.PASTE_COLUMN_WIDTHS;
		} else if ( i3label.equals(Labels.getLabel("paste.comment")) ) {
			return Range.PASTE_COMMENTS;
		} else if ( i3label.equals(Labels.getLabel("paste.formula")) ) {
			return Range.PASTE_FORMULAS;
		} else if ( i3label.equals(Labels.getLabel("paste.formulaWithNumFmt")) ) {
			return Range.PASTE_FORMULAS_AND_NUMBER_FORMATS;
		} else if ( i3label.equals(Labels.getLabel("paste.value")) ) {
			return Range.PASTE_VALUES;
		} else if ( i3label.equals(Labels.getLabel("paste.valueWithNumFmt")) ) {
			return Range.PASTE_VALUES_AND_NUMBER_FORMATS;
		} else if ( i3label.equals(Labels.getLabel("paste.format")) ) {
			return Range.PASTE_FORMATS;
		} else if ( i3label.equals(Labels.getLabel("paste.validation")) ) {
			return Range.PASTE_VALIDATAION;
		}
		
		return Range.PASTE_ALL;
	}
	
	
	public static int getDefaultPasteOperation() {
		return Range.PASTEOP_NONE;
	}
	
	/**
	 * Returns the paste operation base on i3-label, if no match
	 * <p> Default: returns {@link #Range.PASTEOP_NONE}, if no match 
	 * @param i3label
	 * @return
	 */
	public static int getPasteOperation(String i3label) {
		if (i3label == null || i3label.equals(Labels.getLabel("pasteop.none")) )
			return Range.PASTEOP_NONE;
		if ( i3label.equals(Labels.getLabel("pasteop.add")) ) {
			return Range.PASTEOP_ADD;
		} else if ( i3label.equals(Labels.getLabel("pasteop.sub")) ) {
			return Range.PASTEOP_SUB;
		} else if ( i3label.equals(Labels.getLabel("pasteop.mul")) ) {
			return Range.PASTEOP_MUL;
		} else if ( i3label.equals(Labels.getLabel("pasteop.divide")) ) {
			return Range.PASTEOP_DIV;
		}
		return Range.PASTEOP_NONE;
	}
	
	/**
	 * Returns whether transpose or not
	 * <p> Default: returns false
	 * @param i3label
	 * @return 
	 */
	public static boolean isTranspose(String i3label) {
		if (i3label == null || !i3label.equals(Labels.getLabel("paste.transpose")))
			return false;
		return true;
	}
	
	public static void createpPasteSpecialDialog(Spreadsheet spreadsheet, Component parent) {
		Executions.createComponents(
				Consts._PasteSpecialDialog_zul, 
				parent, 
				Zssapps.newSpreadsheetArg(spreadsheet));
	}
}