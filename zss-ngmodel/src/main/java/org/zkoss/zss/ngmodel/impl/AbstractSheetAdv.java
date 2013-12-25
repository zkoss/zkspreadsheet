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

import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NSheet;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class AbstractSheetAdv implements NSheet,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract AbstractRowAdv getRow(int rowIdx, boolean proxy);
	/*package*/ abstract AbstractRowAdv getOrCreateRow(int rowIdx);
//	/*package*/ abstract int getRowIndex(AbstractRowAdv row);
	
	/*package*/ abstract NColumn getColumn(int columnIdx, boolean proxy);
	/*package*/ abstract AbstractColumnArrayAdv getOrSplitColumnArray(int index);
	
//	/*package*/ abstract ColumnAdv getOrCreateColumn(int columnIdx);
//	/*package*/ abstract int getColumnIndex(ColumnAdv column);
	
	/*package*/ abstract AbstractCellAdv getCell(int rowIdx, int columnIdx, boolean proxy);
	/*package*/ abstract AbstractCellAdv getOrCreateCell(int rowIdx, int columnIdx);
	
	
	/*package*/ abstract void copyTo(AbstractSheetAdv sheet);
	/*package*/ abstract void setSheetName(String name);
	
	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
	
	/*package*/ ModelInternalEvent createModelInternalEvent(String name, Object... data){
		return ModelInternalEvents.createModelInternalEvent(name,this,data);
	}
	
}
