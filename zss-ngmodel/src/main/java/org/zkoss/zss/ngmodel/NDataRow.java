package org.zkoss.zss.ngmodel;

import java.util.Iterator;

public interface NDataRow {

	public Iterator<NDataCell> getDataCellIterator();
	public int getIndex();
}
