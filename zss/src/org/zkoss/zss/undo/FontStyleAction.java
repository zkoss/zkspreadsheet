/* FontStyleAction.java

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
import org.zkoss.zss.api.CellOperationUtil.FontStyleApplier;
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
public class FontStyleAction extends AbstractUndoableAction {

	private CellStyle[][] _oldStyles = null;
	private CellStyle[][] _newStyles = null;
	
	private final FontStyleApplier _fontStyleApplier;
	
	
	public FontStyleAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,FontStyleApplier styleApplier){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._fontStyleApplier = styleApplier;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old style
		_oldStyles = new CellStyle[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				_oldStyles[i-_row][j-_column] = r.getCellStyle();
			}
		}
		
		if(_newStyles!=null){//reuse the style
			for(int i=_row;i<=_lastRow;i++){
				for(int j=_column;j<=_lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					r.setCellStyle(_newStyles[i-_row][j-_column]);
				}
			}
			_newStyles = null;
		}else{
			Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
			CellOperationUtil.applyFontStyle(r, _fontStyleApplier);
		}
	}

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
		
		//keep last new style, so if redo-again, we will reuse it.
		_newStyles = new CellStyle[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				_newStyles[i-_row][j-_column] = r.getCellStyle();
			}
		}
		
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				r.setCellStyle(_oldStyles[i-_row][j-_column]);
			}
		}
		_oldStyles = null;
	}
}
