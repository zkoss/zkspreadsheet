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
import org.zkoss.zss.undo.imple.AbstractCellDataStyleAction;

/**
 * This undoable action doesn't handle merge cell when undo redo
 * @author dennis
 * 
 */
public class CellPasteAction extends AbstractCellDataStyleAction {

	protected final int _toRow,_toColumn,_toLastRow,_toLastColumn;
	protected final Sheet _toSheet;
	protected final PasteType _pasteType;
	protected final PasteOperation _pasteOperation;
	protected final boolean _skipBlank;
	protected final boolean _transpose;
	public CellPasteAction(String label, 
			Sheet sheet, int row, int column,int lastRow, int lastColumn, 
			Sheet toSheet, int toRow, int toColumn,int toLastRow, int toLastColumn, 
			PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose) {
		super(label, sheet, row, column, lastRow, lastColumn,ReserveType.ALL);
		
		this._toRow = toRow;
		if(toLastRow-toRow < lastRow-row){
			//enlarge the last to same size
			this._toLastRow = _toRow+lastRow-row;
		}else{
			this._toLastRow = toLastRow;
		}
		this._toColumn = toColumn;
		if(toLastColumn-toColumn < lastColumn-column){
			//enlarge the last to same size
			this._toLastColumn = _toColumn+lastColumn-column;
		}else{
			this._toLastColumn = toLastColumn;
		}
		this._toSheet = toSheet;
		this._pasteType = pasteType;
		this._pasteOperation = pasteOperation;
		this._skipBlank = skipBlank;
		this._transpose = transpose;
		
		

	}

	protected int getReservedRow(){
		return _toRow;
	}
	protected int getReservedColumn(){
		return _toColumn;
	}
	protected int getReservedLastRow(){
		return _toLastRow;
	}
	protected int getReservedLastColumn(){
		return _toLastColumn;
	}
	protected Sheet getReservedSheet(){
		return _toSheet;
	}
	
	public Rect getUndoSelection(){
		return new Rect(_toColumn,_toRow,_toLastColumn,_toLastRow);
	}
	public Rect getRedoSelection(){
		return new Rect(_toColumn,_toRow,_toLastColumn,_toLastRow);
	}
	
	//TODO handle merge, unmerge
	protected void applyAction() {
		Range src = Ranges.range(_sheet, _row, _column, _lastRow, _lastColumn);
		Range dest = Ranges.range(_toSheet, _toRow, _toColumn, _toLastRow, _toLastColumn);
		CellOperationUtil.pasteSpecial(src, dest, _pasteType, _pasteOperation, _skipBlank, _transpose);
	}

}
