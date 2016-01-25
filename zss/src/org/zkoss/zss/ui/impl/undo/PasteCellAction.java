/* PasteCellAction.java

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
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.CellSelectionType;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.range.impl.PasteRangeImpl;

/**
 * 
 * @author dennis
 * 
 */
public class PasteCellAction extends AbstractCellDataStyleAction {
	private static final long serialVersionUID = 603418025542199316L;
	
	protected final int _destRow,_destColumn,_destLastRow,_destLastColumn;
	protected final int _reservedDestLastRow,_reservedDestLastColumn;
	protected final Sheet _destSheet;
	protected final boolean _destWholeColumn; //ZSS-717
	
	private Range _pastedRange;
	
//	private final int rlastRow;
//	private final int rlastColumn;
	//ZSS-717
	//@since 3.8.3
	public PasteCellAction(String label, 
			Sheet sheet, int srcRow, int srcColumn,int srcLastRow, int srcLastColumn, boolean srcWholeColumn,
			Sheet destSheet, int destRow, int destColumn,int destLastRow, int destLastColumn, boolean destWholeColumn) {
		super(label, sheet, srcRow, srcColumn, srcLastRow, srcLastColumn,srcWholeColumn,RESERVE_ALL);
		
		this._destRow = destRow;
		this._destColumn = destColumn;
		this._destLastRow = destLastRow;
		this._destLastColumn = destLastColumn;
		
		int srcColNum = srcLastColumn-srcColumn;
		int srcRowNum = srcLastRow-srcRow;
		//enlarge and transpose
		int destWidth = Math.max(destLastColumn-destColumn, srcColNum);
		int destHeight = Math.max(destLastRow-destRow, srcRowNum);
		
		_reservedDestLastRow = _destRow + destHeight;
		_reservedDestLastColumn = _destColumn + destWidth;
		
		this._destSheet = destSheet;
		this._destWholeColumn = destWholeColumn; //ZSS-717
	}
	@Deprecated
	public PasteCellAction(String label, 
			Sheet sheet, int srcRow, int srcColumn,int srcLastRow, int srcLastColumn, 
			Sheet destSheet, int destRow, int destColumn,int destLastRow, int destLastColumn) {
		this(label, 
			sheet, srcRow, srcColumn, srcLastRow, srcLastColumn, false, 
			destSheet, destRow, destColumn, destLastRow, destLastColumn, false);
	}
	
	@Override
	protected boolean isSheetProtected(){
		try{
			return _destSheet.isProtected();
		}catch(Exception x){}
		return true;
	}
	
	@Override
	public Sheet getUndoSheet(){
		return _destSheet;
	}
	@Override
	public Sheet getRedoSheet(){
		return _destSheet;
	}

	@Override
	protected int getReservedRow(){
		return _destRow;
	}
	@Override
	protected int getReservedColumn(){
		return _destColumn;
	}
	@Override
	protected int getReservedLastRow(){
		return _reservedDestLastRow;
	}
	@Override
	protected int getReservedLastColumn(){
		return _reservedDestLastColumn;
	}
	protected Sheet getReservedSheet(){
		return _destSheet;
	}
	@Override
	public AreaRef getUndoSelection(){
		return _pastedRange==null?new AreaRef(_destRow,_destColumn,_destLastRow,_destLastColumn):
			new AreaRef(_pastedRange.getRow(),_pastedRange.getColumn(),_pastedRange.getLastRow(),_pastedRange.getLastColumn());
	}
	@Override
	public AreaRef getRedoSelection(){
		return _pastedRange==null?new AreaRef(_destRow,_destColumn,_destLastRow,_destLastColumn):
			new AreaRef(_pastedRange.getRow(),_pastedRange.getColumn(),_pastedRange.getLastRow(),_pastedRange.getLastColumn());
	}
	
	protected void applyAction() {
		//ZSS-717
		Range src = new RangeImpl(new PasteRangeImpl(_sheet.getInternalSheet(), _row, _column, _lastRow, _lastColumn, false, _wholeColumn), _sheet);
		Range dest = new RangeImpl(new PasteRangeImpl(_destSheet.getInternalSheet(), _destRow, _destColumn, _destLastRow, _destLastColumn, false, _destWholeColumn), _destSheet);
		_pastedRange = CellOperationUtil.paste(src, dest);
		
		CellOperationUtil.fitFontHeightPoints(Ranges.range(_destSheet, dest.getRow(), dest.getColumn(),
				dest.getRow() + (_lastRow - _row), dest.getColumn() + (_lastColumn - _column)));
			
	}

}
