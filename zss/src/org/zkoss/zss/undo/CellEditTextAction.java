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
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.undo.imple.AbstractUndoableAction;
/**
 * 
 * @author dennis
 *
 */
public class CellEditTextAction extends AbstractUndoableAction {

	private String[][] _oldEditTexts = null;
	
	private final String _editText;
	private final String[][] _editTexts;
	
	public CellEditTextAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,String editText){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._editText = editText;
		_editTexts = null;
	}
	public CellEditTextAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,String[][] editTexts){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._editTexts = editTexts;
		_editText = null;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old text
		_oldEditTexts = new String[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				
				CellData data = r.getCellData();
				if(data.isBlank()){
					_oldEditTexts[i-_row][j-_column] = null;
				}else{
					_oldEditTexts[i-_row][j-_column] = r.getCellEditText();
				}
				
				if(_editTexts!=null){
					try{
						r.setCellEditText(_editTexts[i][j]);
					}catch(IllegalFormulaException x){};//eat in this mode
				}
			}
		}
		if(_editText!=null){
			Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
			try{
				r.setCellEditText(_editText);
			}catch(IllegalFormulaException x){};//eat in this mode
		}
		
	}

	@Override
	public boolean isUndoable() {
		return _oldEditTexts!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return _oldEditTexts==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		if(_oldEditTexts!=null){
			for(int i=_row;i<=_lastRow;i++){
				for(int j=_column;j<=_lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					if(_oldEditTexts[i-_row][j-_column]==null){
						r.clearContents();
					}else{
						r.setCellEditText(_oldEditTexts[i-_row][j-_column]);
					}
				}
			}
			_oldEditTexts = null;
		}
	}
}
