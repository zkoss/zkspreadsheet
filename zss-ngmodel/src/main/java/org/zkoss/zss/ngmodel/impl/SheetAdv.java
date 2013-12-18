/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NSheet;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class SheetAdv implements NSheet,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract RowAdv getRow(int rowIdx, boolean proxy);
	/*package*/ abstract RowAdv getOrCreateRow(int rowIdx);
	/*package*/ abstract int getRowIndex(RowAdv row);
	
	/*package*/ abstract ColumnAdv getColumn(int columnIdx, boolean proxy);
	/*package*/ abstract ColumnArrayAdv getOrSplitColumnArray(int index);
	
//	/*package*/ abstract ColumnAdv getOrCreateColumn(int columnIdx);
//	/*package*/ abstract int getColumnIndex(ColumnAdv column);
	
	/*package*/ abstract CellAdv getCell(int rowIdx, int columnIdx, boolean proxy);
	/*package*/ abstract CellAdv getOrCreateCell(int rowIdx, int columnIdx);
	
	
	/*package*/ abstract void copyTo(SheetAdv sheet);
	/*package*/ abstract void setSheetName(String name);
	
	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
	
	ModelInternalEvent createModelInternalEvent(String name, Object... data){
		return ModelInternalEvents.createModelInternalEvent(name,this,data);
	}
	
}
