package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NSheet;

public abstract class SheetAdv implements NSheet,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract RowAdv getRowAt(int rowIdx, boolean proxy);
	/*package*/ abstract RowAdv getOrCreateRowAt(int rowIdx);
	/*package*/ abstract int getRowIndex(RowAdv row);
	
	/*package*/ abstract ColumnAdv getColumnAt(int columnIdx, boolean proxy);
	/*package*/ abstract ColumnAdv getOrCreateColumnAt(int columnIdx);
	/*package*/ abstract int getColumnIndex(ColumnAdv column);
	
	/*package*/ abstract CellAdv getCellAt(int rowIdx, int columnIdx, boolean proxy);
	/*package*/ abstract CellAdv getOrCreateCellAt(int rowIdx, int columnIdx);
	
	
	/*package*/ abstract void copyTo(SheetAdv sheet);
	/*package*/ abstract void setSheetName(String name);
	
	/*package*/ abstract void onModelEvent(ModelEvent event);
}
