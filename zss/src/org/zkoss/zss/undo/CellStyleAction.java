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
package org.zkoss.zss.undo;


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
public class CellStyleAction extends AbstractUndoableAction {

	private CellStyle[][] oldStyles = null;
	private CellStyle[][] newStyles = null;
	
	private final CellStyleApplier styleApplier;
	
	
	public CellStyleAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,CellStyleApplier styleApplier){
		super(label,sheet,row,column,lastRow,lastColumn);
		this.styleApplier = styleApplier;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old style
		oldStyles = new CellStyle[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				oldStyles[i-_row][j-_column] = r.getCellStyle();
			}
		}
		
		if(newStyles!=null){//reuse the style
			for(int i=_row;i<=_lastRow;i++){
				for(int j=_column;j<=_lastColumn;j++){
					Range r = Ranges.range(_sheet,i,j);
					r.setCellStyle(newStyles[i-_row][j-_column]);
				}
			}
			newStyles = null;
		}else{
			Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
			CellOperationUtil.applyCellStyle(r, styleApplier);
		}
	}

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
		
		//keep last new style, so if redo-again, we will reuse it.
		newStyles = new CellStyle[_lastRow-_row+1][_lastColumn-_column+1];
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				newStyles[i-_row][j-_column] = r.getCellStyle();
			}
		}
		
		for(int i=_row;i<=_lastRow;i++){
			for(int j=_column;j<=_lastColumn;j++){
				Range r = Ranges.range(_sheet,i,j);
				r.setCellStyle(oldStyles[i-_row][j-_column]);
			}
		}
		oldStyles = null;
	}

}
