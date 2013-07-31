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
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.undo.impl.AbstractUndoableAction;
import org.zkoss.zss.undo.impl.ReserveUtil;
import org.zkoss.zss.undo.impl.ReserveUtil.ReservedResult;

/**
 * This undoable action doesn't handle merge cell when undo redo yet
 * @author dennis
 * 
 */
public class CutCellAction extends AbstractUndoableAction {

	
//	private CellStyle[][] _destNewStyles = null;
//	private ReservedCellData[][] _destNewData = null;
	private ReservedResult _destOldReserve = null;

	private ReservedResult _srcOldReserve = null;
	
	protected final Sheet _destSheet;
	protected final int _destRow,_destColumn,_destLastRow,_destLastColumn;
	protected final int _reservedDestLastRow,_reservedDestLastColumn;
	private Range _pastedRange;
	public CutCellAction(String label,Sheet sheet,int srcRow, int srcColumn, int srcLastRow,int srcLastColumn,
			Sheet destSheet,int destRow, int destColumn, int destLastRow,int destLastColumn){
		super(label,sheet,srcRow,srcColumn,srcLastRow,srcLastColumn);
		_destSheet = destSheet;
		_destRow = destRow;
		_destColumn = destColumn;
		_destLastRow = destLastRow;
		_destLastColumn = destLastColumn;
		
		_reservedDestLastRow = _destRow + Math.max(destLastRow-destRow, srcLastRow-srcRow);
		_reservedDestLastColumn = _destColumn + Math.max(destLastColumn-destColumn, srcLastColumn-srcColumn);
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old style/data of src and dest
		
		_srcOldReserve = ReserveUtil.reserve(_sheet, _row, _column, _lastRow, _lastColumn, ReserveUtil.ReserveType.ALL);
		_destOldReserve = ReserveUtil.reserve(_destSheet, _destRow, _destColumn, _reservedDestLastRow, _reservedDestLastColumn, ReserveUtil.ReserveType.ALL);
		
		applyAction();
	}

	
	protected void applyAction(){
		Range src = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
		Range dest = Ranges.range(_destSheet,_destRow,_destColumn,_destLastRow,_destLastColumn);
		_pastedRange = CellOperationUtil.cut(src, dest);
	}
	
	@Override
	public boolean isUndoable() {
		return _destOldReserve!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return _destOldReserve==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		
		ReserveUtil.restore(_destOldReserve);
		ReserveUtil.restore(_srcOldReserve);
		_srcOldReserve = null;
		_destOldReserve = null;
	}
	
	@Override
	public Rect getUndoSelection(){
		return _pastedRange==null?new Rect(_destColumn,_destRow,_destLastColumn,_destLastRow):
			new Rect(_pastedRange.getColumn(),_pastedRange.getRow(),_pastedRange.getLastColumn(),_pastedRange.getLastRow());
	}
	@Override
	public Rect getRedoSelection(){
		return _pastedRange==null?new Rect(_destColumn,_destRow,_destLastColumn,_destLastRow):
			new Rect(_pastedRange.getColumn(),_pastedRange.getRow(),_pastedRange.getLastColumn(),_pastedRange.getLastRow());
	}

}
