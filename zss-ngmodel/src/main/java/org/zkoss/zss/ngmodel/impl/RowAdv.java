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

import org.zkoss.zss.ngmodel.NRow;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class RowAdv implements NRow,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract CellAdv getCell(int columnIdx, boolean proxy);

	/*package*/ abstract CellAdv getOrCreateCell(int columnIdx);
	
	/*package*/ abstract void onModelEvent(ModelInternalEvent event);

	/*package*/ abstract void clearCell(int start, int end);

	/*package*/ abstract void insertCell(int start, int size);

	/*package*/ abstract void deleteCell(int start, int size);

	/*package*/ abstract int getCellIndex(CellAdv cell);
}
