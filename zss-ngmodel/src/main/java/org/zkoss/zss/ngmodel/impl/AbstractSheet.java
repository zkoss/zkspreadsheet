package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NSheet;

public abstract class AbstractSheet implements NSheet{

	abstract AbstractRow getRowAt(int rowIdx, boolean proxy);
	abstract AbstractRow getOrCreateRowAt(int rowIdx);
	abstract int getRowIndex(AbstractRow row);
	
	abstract AbstractColumn getColumnAt(int columnIdx, boolean proxy);
	abstract AbstractColumn getOrCreateColumnAt(int columnIdx);
	abstract int getColumnIndex(AbstractColumn column);
	
	abstract AbstractCell getCellAt(int rowIdx, int columnIdx, boolean proxy);
	abstract AbstractCell getOrCreateCellAt(int rowIdx, int columnIdx);
	
	
	abstract void copyTo(AbstractSheet sheet);
	abstract void setSheetName(String name);
	
	void release() {
	}
	void onModelEvent(ModelEvent event) {
	}
}
