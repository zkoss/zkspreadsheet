/* PasteHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/8/3 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl.ua;

import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.AreaRefWithType;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.range.impl.PasteRangeImpl;
import org.zkoss.zss.ui.CellSelectionType;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.UserActionContext.Clipboard;
import org.zkoss.zss.ui.impl.undo.CutCellAction;
import org.zkoss.zss.ui.impl.undo.PasteCellAction;
import org.zkoss.zss.ui.sys.UndoableActionManager;

public class PasteHandler extends AbstractHandler {
	private static final long serialVersionUID = -6262315007795949652L;

	@Override
	protected boolean processAction(UserActionContext ctx) {
		Clipboard cb = ctx.getClipboard();
		if(cb==null){
			//don't handle it , so give a chance to do client paste
			return false;
		}
		
		Book book = ctx.getBook();
		Sheet destSheet = ctx.getSheet();
		Sheet srcSheet = cb.getSheet();
		if(srcSheet==null){
			showInfoMessage(Labels.getLabel("zss.actionhandler.msg.cant_find_sheet_to_paste"));
			ctx.clearClipboard();
			return true;
		}
		//ZSS-717
		// 20160125, henrichen: The spec is we copy only those values within 
		// MaxVisible range so we save some time 
		AreaRefWithType src = cb.getSelectionWithType();
		AreaRefWithType selection = ctx.getSelectionWithType();
		CellSelectionType srcType = src.getSelType();
		CellSelectionType dstType = selection.getSelType();
		final boolean srcWholeRow = srcType == CellSelectionType.ALL || srcType == CellSelectionType.ROW;
		final boolean srcWholeColumn = srcType == CellSelectionType.ALL || srcType == CellSelectionType.COLUMN;
		final boolean dstWholeRow = dstType == CellSelectionType.ALL || dstType == CellSelectionType.ROW;
		final boolean dstWholeColumn = dstType == CellSelectionType.ALL || dstType == CellSelectionType.COLUMN;
				
		int srcLastRow = srcWholeColumn ? 
				Math.max(cb.getSheetMaxVisibleRows() - 1, src.getLastRow()) : src.getLastRow();
		int srcLastCol = srcWholeRow ?
				Math.max(cb.getSheetMaxVisibleColumns() - 1, src.getLastColumn()) : src.getLastColumn();
		int dstLastRow = dstWholeColumn ?
				Math.max(ctx.getSheetMaxVisibleRows() - 1, selection.getLastRow()) : selection.getLastRow();
		int dstLastCol = dstWholeRow ?
				Math.max(ctx.getSheetMaxVisibleColumns() - 1,  selection.getLastColumn()) : selection.getLastColumn();
				
		if (srcWholeColumn && dstWholeColumn && srcLastRow != dstLastRow) {
			dstLastRow = srcLastRow = Math.max(dstLastRow, srcLastRow);
		}
		if (srcWholeRow && dstWholeRow && srcLastCol != dstLastCol) {
			dstLastCol = srcLastCol = Math.max(dstLastCol, srcLastCol);
		}
				
		Range srcRange = new RangeImpl(new PasteRangeImpl(srcSheet.getInternalSheet(), src.getRow(),
				src.getColumn(), srcLastRow,srcLastCol, srcWholeRow, srcWholeColumn), srcSheet);

		Range destRange = new RangeImpl(new PasteRangeImpl(destSheet.getInternalSheet(), selection.getRow(),
				selection.getColumn(), dstLastRow, dstLastCol, dstWholeRow, dstWholeColumn), destSheet);
		
		if (destRange.isProtected()) {
			showProtectMessage();
			return true;
		} else if (cb.isCutMode() && srcRange.isProtected()) {
			showProtectMessage();
			return true;
		}
		
		UndoableActionManager uam = ctx.getSpreadsheet().getUndoableActionManager();
		
		if(cb.isCutMode()){
			uam.doAction(new CutCellAction(Labels.getLabel("zss.undo.cut"),
				srcSheet, srcRange.getRow(), srcRange.getColumn(),srcRange.getLastRow(), srcRange.getLastColumn(),srcRange.isWholeColumn(),srcRange.isWholeRow(), //ZSS-1277
				destSheet, destRange.getRow(), destRange.getColumn(),destRange.getLastRow(), destRange.getLastColumn(),destRange.isWholeColumn(),destRange.isWholeRow())); //ZSS-1277
			ctx.clearClipboard();
		}else{
			uam.doAction(new PasteCellAction(Labels.getLabel("zss.undo.paste"),
				srcSheet, srcRange.getRow(), srcRange.getColumn(),srcRange.getLastRow(), srcRange.getLastColumn(),srcRange.isWholeColumn(),srcRange.isWholeRow(), //ZSS-1277
				destSheet, destRange.getRow(), destRange.getColumn(),destRange.getLastRow(), destRange.getLastColumn(),destRange.isWholeColumn(),destRange.isWholeRow())); //ZSS-1277
			// ZSS-1380 clear highlight when pasting range overlaps the source range
			if (srcSheet.equals(destSheet) && src.overlap(selection)){
				ctx.clearClipboard();
			}

		}
		return true;
	}

}
