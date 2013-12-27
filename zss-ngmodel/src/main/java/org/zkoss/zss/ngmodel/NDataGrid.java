package org.zkoss.zss.ngmodel;

import java.util.Iterator;

/**
 * data grid to store cell data.
 * @author dennis
 *
 */
public interface NDataGrid {

	//basic storage method
	public NCellValue getValue(int row, int column);
	public void setValue(int row, int column, NCellValue value);
	public boolean validateValue(int row, int column, NCellValue value);
	
	//support operations
	public boolean supportOperations();
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int rowIdx, int size);
	public void deleteColumn(int rowIdx, int size);
	
	//support data iterator
	public boolean supportDataIterator();
	public Iterator<NDataRow> getRowIterator();
	public Iterator<NDataCell> getCellIterator(int row);
	
}
