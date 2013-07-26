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


import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;
/**
 * 
 * @author dennis
 *
 */
public abstract class AbstractCellDataAction extends AbstractUndoableAction {

	private String[][] oldTexts = null;
	
	
	public AbstractCellDataAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn){
		super(label,sheet,row,column,lastRow,lastColumn);
	}
	
	protected Sheet getReservedSheet(){
		return _sheet;
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
		Sheet sheet = getReservedSheet();
		oldTexts = new String[lastRow-row+1][lastColumn-column+1];
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				CellData d = r.getCellData();
				if(d.isBlank()){
					oldTexts[i-row][j-column] = null;
				}else{
					oldTexts[i-row][j-column] = d.getEditText();
				}
			}
		}
		
		applyAction();
	}

	
	protected abstract void applyAction();
	
	@Override
	public boolean isUndoable() {
		return oldTexts!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return oldTexts==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		
		int row = getReservedRow();
		int column = getReservedColumn();
		int lastRow = getReservedLastRow();
		int lastColumn = getReservedLastColumn();
		Sheet sheet = getReservedSheet();
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				try{
					r.setCellEditText(oldTexts[i-row][j-column]);
				}catch(IllegalFormulaException x){};//eat in this mode
			}
		}
		oldTexts = null;
	}

}
