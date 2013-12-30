package org.zkoss.zss.ngmodel;

import java.util.Iterator;

/**
 * data grid to store cell data.
 * @author dennis
 *
 */
public interface NDataGrid {

	//basic storage method
	public int getStartCellIndex(int rowIdx);
	public int getEndCellIndex(int rowIdx);
	
	public NCellValue getValue(int row, int column);
	public void setValue(int row, int column, NCellValue value);
	public boolean validateValue(int row, int column, NCellValue value);
	
	//support operations
	public boolean isSupportedOperations();
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int rowIdx, int size);
	public void deleteColumn(int rowIdx, int size);
	
	//support data iterator
	public boolean isSupportedDataIterator();
	public Iterator<NDataRow> getRowIterator();
	public Iterator<NDataCell> getCellIterator(int row);
	
	//support start end index
	public boolean isSupportedDataStartEndIndex();
	public int getStartRowIndex();
	public int getEndRowIndex();
}
