/* ClearCellAction.java

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
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
/**
 * 
 * @author dennis
 *
 */
public class ClearCellAction extends AbstractUndoableAction {

	
	public enum Type{
		CONTENT,
		STYLE,
		ALL
	}
	
	private CellStyle[][] oldStyles = null;
	private String[][] oldEditTexts = null;
	private Type _type;
	
	public ClearCellAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,Type type){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._type = type;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		/*
		 * Refer to BookHelper#clearCell. it only clear the formula and set the value to string null
		 */
		//keep old text
		boolean clearContent = _type==Type.CONTENT||_type==Type.ALL;
		boolean clearStyle = _type==Type.STYLE||_type==Type.ALL;
		
		if(clearContent){
			oldEditTexts = new String[_lastRow-_row+1][_lastColumn-_column+1];
		}
		if(clearStyle){
			oldStyles = new CellStyle[_lastRow-_row+1][_lastColumn-_column+1];
		}
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				if(clearContent){
					CellData data = r.getCellData();
					if(data.isBlank()){
						oldEditTexts[i-_row][j-_column] = null;
					}else{
						oldEditTexts[i-_row][j-_column] = r.getCellEditText();
					}
				}
				if(clearStyle){
					oldStyles[i-_row][j-_column] = r.getCellStyle();
				}
				
			}
		}
		Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
		switch(_type){
		case CONTENT:
			CellOperationUtil.clearContents(r);
			break;
		case STYLE:
			CellOperationUtil.clearStyles(r);
			break;
		case ALL:
			CellOperationUtil.clearAll(r);
			break;
		}
	}

	@Override
	public boolean isUndoable() {
		return (oldEditTexts!=null||oldStyles!=null) && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return (oldEditTexts==null&&oldStyles==null) && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		if(oldEditTexts!=null){
			for(int i=_row;i<=_lastRow;i++){
				for(int j=_column;j<=_lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					if(oldEditTexts[i-_row][j-_column]==null){
//						r.clearContents(); //no need to do anything, it is clear
					}else{
						r.setCellEditText(oldEditTexts[i-_row][j-_column]);
					}
				}
			}
			oldEditTexts = null;
		}
		
		if(oldStyles!=null){
			for(int i=_row;i<=_lastRow;i++){
				for(int j=_column;j<=_lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					r.setCellStyle(oldStyles[i-_row][j-_column]);
				}
			}
			oldStyles = null;
		}

	}
}
