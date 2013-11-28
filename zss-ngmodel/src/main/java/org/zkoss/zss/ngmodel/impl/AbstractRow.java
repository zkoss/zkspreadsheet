package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NRow;

public abstract class AbstractRow implements NRow,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;

	/*package*/ abstract AbstractCell getCellAt(int columnIdx, boolean proxy);

	/*package*/ abstract AbstractCell getOrCreateCellAt(int columnIdx);
	
	/*package*/ abstract void onModelEvent(ModelEvent event);

	/*package*/ abstract void clearCell(int start, int end);

	/*package*/ abstract void insertCell(int start, int size);

	/*package*/ abstract void deleteCell(int start, int size);

	/*package*/ abstract int getCellIndex(AbstractCell cell);
}
