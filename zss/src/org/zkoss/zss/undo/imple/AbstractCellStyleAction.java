/* AbstractCellStyleAction.java

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
package org.zkoss.zss.undo.imple;


import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.CellOperationUtil.CellStyleApplier;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.undo.imple.AbstractUndoableAction;
/**
 * 
 * @author dennis
 *
 */
public abstract class AbstractCellStyleAction extends AbstractUndoableAction {

	private CellStyle[][] oldStyles = null;
	private CellStyle[][] newStyles = null;
	
	
	public AbstractCellStyleAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn){
		super(label,sheet,row,column,lastRow,lastColumn);
	}
	
	protected int getReservedRow(){
		return _row;
	}
	protected int getReservedColumn(){
		return _column;
	}
	protected int getReservedLastRow(){
		return _lastRow;
	}
	protected int getReservedLastColumn(){
		return _lastColumn;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old style

		int row = getReservedRow();
		int column = getReservedColumn();
		int lastRow = getReservedLastRow();
		int lastColumn = getReservedLastColumn();
		oldStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				oldStyles[i-row][j-column] = r.getCellStyle();
			}
		}
		
		if(newStyles!=null){//reuse the style
			for(int i=row;i<=lastRow;i++){
				for(int j=column;j<=lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					r.setCellStyle(newStyles[i-row][j-column]);
				}
			}
			newStyles = null;
		}else{
			applyAction();
		}
	}

	
	protected abstract void applyAction();
	
	@Override
	public boolean isUndoable() {
		return oldStyles!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return oldStyles==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		
		int row = getReservedRow();
		int column = getReservedColumn();
		int lastRow = getReservedLastRow();
		int lastColumn = getReservedLastColumn();
		//keep last new style, so if redo-again, we will reuse it.
		newStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				newStyles[i-row][j-column] = r.getCellStyle();
			}
		}
		
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				r.setCellStyle(oldStyles[i-row][j-column]);
			}
		}
		oldStyles = null;
	}

}
