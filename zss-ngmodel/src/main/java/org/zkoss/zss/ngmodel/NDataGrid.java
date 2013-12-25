package org.zkoss.zss.ngmodel;

import java.util.Iterator;

/**
 * data grid to store cell data.
 * @author dennis
 *
 */
public interface NDataGrid {

	//basic storage method?
	public Object getValue(int row, int column);
	public void setValue(int row, int column, Object value);
	
	
	//support operations
	public boolean supportOperations();
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int rowIdx, int size);
	public void deleteColumn(int rowIdx, int size);
	
	//
	public boolean supportDataIterator();
	public Iterator<NDataRow> getDataRowIterator();
	
//	//TODO
	public boolean validateValue(int row, int column, Object value);
}
