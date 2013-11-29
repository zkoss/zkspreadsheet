package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NSheet;

public abstract class SheetAdv implements NSheet,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract RowAdv getRow(int rowIdx, boolean proxy);
	/*package*/ abstract RowAdv getOrCreateRow(int rowIdx);
	/*package*/ abstract int getRowIndex(RowAdv row);
	
	/*package*/ abstract ColumnAdv getColumn(int columnIdx, boolean proxy);
	/*package*/ abstract ColumnAdv getOrCreateColumn(int columnIdx);
	/*package*/ abstract int getColumnIndex(ColumnAdv column);
	
	/*package*/ abstract CellAdv getCell(int rowIdx, int columnIdx, boolean proxy);
	/*package*/ abstract CellAdv getOrCreateCell(int rowIdx, int columnIdx);
	
	
	/*package*/ abstract void copyTo(SheetAdv sheet);
	/*package*/ abstract void setSheetName(String name);
	
	/*package*/ abstract void onModelEvent(ModelEvent event);
}
