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
package org.zkoss.zss.undo.impl;


import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
/**
 * 
 * @author dennis
 *
 */
public abstract class AbstractCellDataStyleAction extends AbstractUndoableAction {

	private CellStyle[][] _oldStyles = null;
	private CellStyle[][] _newStyles = null;
	private ReservedCellData[][] _oldData = null;
	private ReservedCellData[][] _newData = null;
	
	
	private final ReserveType _reserveType;
	public enum ReserveType {
		DATA,STYLE,ALL
	}
	
	public AbstractCellDataStyleAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,ReserveType reserveType){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._reserveType=reserveType;
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
	protected Sheet getReservedSheet(){
		return _sheet;
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
		
		switch(_reserveType){
		case DATA:
			_oldData = new ReservedCellData[lastRow-row+1][lastColumn-column+1];
			break;
		case STYLE:
			_oldStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
			break;
		case ALL:
			_oldData = new ReservedCellData[lastRow-row+1][lastColumn-column+1];
			_oldStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
		}
		
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				if(_oldStyles!=null){
					_oldStyles[i-row][j-column] = r.getCellStyle();
				}
				if(_oldData!=null){
					_oldData[i-row][j-column] = ReservedCellData.reserve(r);
				}
			}
		}
		
		if(_newStyles!=null || _newData!=null){//reuse the style
			for(int i=row;i<=lastRow;i++){
				for(int j=column;j<=lastColumn;j++){
					Range r = Ranges.range(sheet,i,j);
					if(_newStyles!=null){
						r.setCellStyle(_newStyles[i-row][j-column]);
					}
					if(_newData!=null){
						try{
							_newData[i-row][j-column].apply(r);
						}catch(IllegalFormulaException x){};//eat in this mode
					}
				}
			}
			_newStyles = null;
			_newData = null;
		}else{
			//first time
			applyAction();
		}
	}

	
	protected abstract void applyAction();
	
	@Override
	public boolean isUndoable() {
		return _oldStyles!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return _oldStyles==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		
		int row = getReservedRow();
		int column = getReservedColumn();
		int lastRow = getReservedLastRow();
		int lastColumn = getReservedLastColumn();
		Sheet sheet = getReservedSheet();
		//keep last new style, so if redo-again, we will reuse it.
		
		switch(_reserveType){
		case DATA:
			_newData = new ReservedCellData[lastRow-row+1][lastColumn-column+1];
			break;
		case STYLE:
			_newStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
			break;
		case ALL:
			_newData = new ReservedCellData[lastRow-row+1][lastColumn-column+1];
			_newStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
		}
		
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				if(_newStyles!=null){
					_newStyles[i-row][j-column] = r.getCellStyle();
				}
				
				if(_newData!=null){
					_newData[i-row][j-column] = ReservedCellData.reserve(r);
				}
			}
		}
		
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				if(_oldStyles!=null){
					r.setCellStyle(_oldStyles[i-row][j-column]);
				}
				if(_oldData!=null){
					_oldData[i-row][j-column].apply(r);
				}
			}
		}
		_oldData = null;
		_oldStyles = null;
	}

}
