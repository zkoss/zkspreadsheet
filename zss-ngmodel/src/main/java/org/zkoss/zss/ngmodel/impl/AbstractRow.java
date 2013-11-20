package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NRow;

public abstract class AbstractRow implements NRow{

	abstract AbstractCell getCellAt(int columnIdx, boolean proxy);

	abstract AbstractCell getOrCreateCellAt(int columnIdx);

	void release() {
	}
	
	void onModelEvent(ModelEvent event) {
	}

	abstract void clearCell(int start, int end);

	abstract void insertCell(int start, int size);

	abstract void deleteCell(int start, int size);

	abstract AbstractSheet getSheet();

	abstract int getCellIndex(AbstractCell cell);
}
