package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NSheet;

public abstract class AbstractSheet implements NSheet,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract AbstractRow getRowAt(int rowIdx, boolean proxy);
	/*package*/ abstract AbstractRow getOrCreateRowAt(int rowIdx);
	/*package*/ abstract int getRowIndex(AbstractRow row);
	
	/*package*/ abstract AbstractColumn getColumnAt(int columnIdx, boolean proxy);
	/*package*/ abstract AbstractColumn getOrCreateColumnAt(int columnIdx);
	/*package*/ abstract int getColumnIndex(AbstractColumn column);
	
	/*package*/ abstract AbstractCell getCellAt(int rowIdx, int columnIdx, boolean proxy);
	/*package*/ abstract AbstractCell getOrCreateCellAt(int rowIdx, int columnIdx);
	
	
	/*package*/ abstract void copyTo(AbstractSheet sheet);
	/*package*/ abstract void setSheetName(String name);
	
	/*package*/ abstract void onModelEvent(ModelEvent event);
}
