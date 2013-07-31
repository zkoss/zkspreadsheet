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
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.undo.impl.AbstractCellDataStyleAction;
import org.zkoss.zss.undo.impl.ReserveUtil;
/**
 * 
 * @author dennis
 *
 */
public class FontStyleAction extends AbstractCellDataStyleAction {
	
	private final FontStyleApplier _fontStyleApplier;
	
	
	public FontStyleAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,FontStyleApplier styleApplier){
		super(label,sheet,row,column,lastRow,lastColumn,ReserveUtil.RESERVE_STYLE);
		this._fontStyleApplier = styleApplier;
	}
	
	protected void applyAction(){
		Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
		CellOperationUtil.applyFontStyle(r, _fontStyleApplier);
	}
}
