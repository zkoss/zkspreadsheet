/* CellEditTextAction.java

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
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.undo.imple.AbstractUndoableAction;
/**
 * 
 * @author dennis
 *
 */
public class ClearContentAction extends AbstractUndoableAction {

	private String[][] oldEditTexts = null;
	
	
	public ClearContentAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn){
		super(label,sheet,row,column,lastRow,lastColumn);
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		/*
		 * Refer to BookHelper#clearCell. it only clear the formula and set the value to string null
		 */
		//keep old text
		oldEditTexts = new String[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				
				CellData data = r.getCellData();
				if(data.isBlank()){
					oldEditTexts[i-_row][j-_column] = null;
				}else{
					oldEditTexts[i-_row][j-_column] = r.getCellEditText();
				}
			}
		}
		Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
		CellOperationUtil.clearContents(r);
	}

	@Override
	public boolean isUndoable() {
		return oldEditTexts!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return oldEditTexts==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		if(oldEditTexts!=null){
			for(int i=_row;i<=_lastRow;i++){
				for(int j=_column;j<=_lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					if(oldEditTexts[i-_row][j-_column]==null){
						r.clearContents();
					}else{
						r.setCellEditText(oldEditTexts[i-_row][j-_column]);
					}
				}
			}
			oldEditTexts = null;
		}
	}
}
