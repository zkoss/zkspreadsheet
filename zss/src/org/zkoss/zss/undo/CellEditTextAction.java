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

import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.undo.imple.AbstractUndoableAction;
/**
 * 
 * @author dennis
 *
 */
public class CellEditTextAction extends AbstractUndoableAction {

	private String[][] oldEditTexts = null;
	
	private final String editText;
	private final String[][] editTexts;
	
	public CellEditTextAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,String editText){
		super(label,sheet,row,column,lastRow,lastColumn);
		this.editText = editText;
		editTexts = null;
	}
	public CellEditTextAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,String[][] editTexts){
		super(label,sheet,row,column,lastRow,lastColumn);
		this.editTexts = editTexts;
		editText = null;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old text
		oldEditTexts = new String[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				oldEditTexts[i-_row][j-_column] = r.getCellEditText();
				if(editTexts!=null){
					try{
						r.setCellEditText(editTexts[i][j]);
					}catch(IllegalFormulaException x){};//eat in this mode
				}
			}
		}
		if(editText!=null){
			Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
			try{
				r.setCellEditText(editText);
			}catch(IllegalFormulaException x){};//eat in this mode
		}
		
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
					r.setCellEditText(oldEditTexts[i-_row][j-_column]);
				}
			}
			oldEditTexts = null;
		}
	}
}
