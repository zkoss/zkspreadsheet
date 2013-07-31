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
import org.zkoss.zss.undo.impl.AbstractUndoableAction;
import org.zkoss.zss.undo.impl.ReserveUtil;
import org.zkoss.zss.undo.impl.ReserveUtil.ReservedResult;
/**
 * 
 * @author dennis
 *
 */
public class MergeCellAction extends AbstractUndoableAction {
	
	private final boolean _accross;
	ReservedResult _oldReserve;
//	ReservedResult _newReserve;
	
	
	public MergeCellAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,boolean accross){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._accross = accross;
	}
	
	protected void applyAction(){
		Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
		CellOperationUtil.merge(r, _accross);
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old style
		
		_oldReserve = ReserveUtil.reserve(_sheet, _row, _column, _lastRow, _lastColumn, ReserveUtil.RESERVE_ALL);
		
		applyAction();
	}
	
	@Override
	public boolean isUndoable() {
		return _oldReserve!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return _oldReserve==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		_oldReserve.restore();
		_oldReserve = null;
	}
}
