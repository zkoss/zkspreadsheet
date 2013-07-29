/* CellStyleAction.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/7/25, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
 */
package org.zkoss.zss.undo;

import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.undo.impl.AbstractCellDataStyleAction;

/**
 * This undoable action doesn't handle merge cell when undo redo
 * @author dennis
 * 
 */
public class PasteCellAction extends AbstractCellDataStyleAction {

	protected final int _toRow,_toColumn,_toLastRow,_toLastColumn;
	protected final int _reservedLastRow,_reservedLastColumn;
	protected final Sheet _toSheet;
	protected final PasteType _pasteType;
	protected final PasteOperation _pasteOperation;
	protected final boolean _skipBlank;
	protected final boolean _transpose;
	
	private Range _pastedRange;
	
//	private final int rlastRow;
//	private final int rlastColumn;
	
	public PasteCellAction(String label, 
			Sheet sheet, int row, int column,int lastRow, int lastColumn, 
			Sheet toSheet, int toRow, int toColumn,int toLastRow, int toLastColumn, 
			PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose) {
		super(label, sheet, row, column, lastRow, lastColumn,ReserveType.ALL);
		this._transpose = transpose;
		
		this._toRow = toRow;
		this._toColumn = toColumn;
		this._toLastRow = toLastRow;
		this._toLastColumn = toLastColumn;
		
		int srcColNum = lastColumn-column;
		int srcRowNum = lastRow-row;

		int destWidth = Math.max(toLastColumn-toColumn, transpose?srcRowNum:srcColNum);
		int destHeight = Math.max(toLastRow-toRow, transpose?srcColNum:srcRowNum);
		
		_reservedLastRow = _toRow + destHeight;
		_reservedLastColumn = _toColumn + destWidth;
		
		this._toSheet = toSheet;
		this._pasteType = pasteType;
		this._pasteOperation = pasteOperation;
		this._skipBlank = skipBlank;
		
	}

	@Override
	protected int getReservedRow(){
		return _toRow;
	}
	@Override
	protected int getReservedColumn(){
		return _toColumn;
	}
	@Override
	protected int getReservedLastRow(){
		return _reservedLastRow;
	}
	@Override
	protected int getReservedLastColumn(){
		return _reservedLastColumn;
	}
	protected Sheet getReservedSheet(){
		return _toSheet;
	}
	@Override
	public Rect getUndoSelection(){
		return _pastedRange==null?new Rect(_toColumn,_toRow,_toLastColumn,_toLastRow):
			new Rect(_pastedRange.getColumn(),_pastedRange.getRow(),_pastedRange.getLastColumn(),_pastedRange.getLastRow());
	}
	@Override
	public Rect getRedoSelection(){
		return _pastedRange==null?new Rect(_toColumn,_toRow,_toLastColumn,_toLastRow):
			new Rect(_pastedRange.getColumn(),_pastedRange.getRow(),_pastedRange.getLastColumn(),_pastedRange.getLastRow());
	}
	
	//TODO handle merge, unmerge
	protected void applyAction() {
		Range src = Ranges.range(_sheet, _row, _column, _lastRow, _lastColumn);
		Range dest = Ranges.range(_toSheet, _toRow, _toColumn, _toLastRow, _toLastColumn);
		_pastedRange = CellOperationUtil.pasteSpecial(src, dest, _pasteType, _pasteOperation, _skipBlank, _transpose);
	}

}
