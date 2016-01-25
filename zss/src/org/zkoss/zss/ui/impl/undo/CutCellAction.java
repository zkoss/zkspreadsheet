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
package org.zkoss.zss.ui.impl.undo;

import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.range.impl.PasteRangeImpl;

/**
 * 
 * @author dennis
 * 
 */
public class CutCellAction extends Abstract2DCellDataStyleAction {
	private static final long serialVersionUID = -7964993492937174380L;
	
	protected final int _reservedDestLastRow,_reservedDestLastColumn;
	private Range _pastedRange;
	
	//ZSS-717
	//@since 3.8.3
	public CutCellAction(String label,Sheet sheet,int srcRow, int srcColumn, int srcLastRow,int srcLastColumn, boolean srcWholeColumn,
			Sheet destSheet,int destRow, int destColumn, int destLastRow,int destLastColumn, boolean destWholeColumn){
		super(label,sheet,srcRow,srcColumn,srcLastRow,srcLastColumn,srcWholeColumn,destSheet,destRow,destColumn,destLastRow,destLastColumn,destWholeColumn,RESERVE_ALL);
		
		//enlarge it, since xrange did
		_reservedDestLastRow = _destRow + Math.max(destLastRow-destRow, srcLastRow-srcRow);
		_reservedDestLastColumn = _destColumn + Math.max(destLastColumn-destColumn, srcLastColumn-srcColumn);
	}
	@Deprecated
	public CutCellAction(String label,Sheet sheet,int srcRow, int srcColumn, int srcLastRow,int srcLastColumn,
			Sheet destSheet,int destRow, int destColumn, int destLastRow,int destLastColumn){
		this(label,sheet,srcRow, srcColumn, srcLastRow,srcLastColumn,false,
			destSheet,destRow, destColumn, destLastRow,destLastColumn,false);
	}

	protected int getReservedDestLastRow(){
		return _reservedDestLastRow;
	}
	protected int getReservedDestLastColumn(){
		return _reservedDestLastColumn;
	}
	
	protected void applyAction(){
		//ZSS-717
		Range src = new RangeImpl(new PasteRangeImpl(_sheet.getInternalSheet(), _row, _column, _lastRow, _lastColumn, false, _wholeColumn), _sheet);
		Range dest = new RangeImpl(new PasteRangeImpl(_destSheet.getInternalSheet(), _destRow, _destColumn, _destLastRow, _destLastColumn, false, _destWholeColumn), _destSheet);
		_pastedRange = CellOperationUtil.cut(src, dest);
		
		CellOperationUtil.fitFontHeightPoints(Ranges.range(_destSheet, dest.getRow(), dest.getColumn(),  
				dest.getRow() + (_lastRow - _row), dest.getColumn() + (_lastColumn - _column)));
	}
		
	@Override
	public AreaRef getUndoSelection(){
		return _pastedRange==null?super.getUndoSelection():
			new AreaRef(_pastedRange.getRow(),_pastedRange.getColumn(),_pastedRange.getLastRow(),_pastedRange.getLastColumn());
	}
	@Override
	public AreaRef getRedoSelection(){
		return _pastedRange==null?super.getRedoSelection():
			new AreaRef(_pastedRange.getRow(),_pastedRange.getColumn(),_pastedRange.getLastRow(),_pastedRange.getLastColumn());
	}

}
