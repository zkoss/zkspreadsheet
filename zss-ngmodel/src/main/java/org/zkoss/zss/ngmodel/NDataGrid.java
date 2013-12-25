package org.zkoss.zss.ngmodel;
/**
 * data grid to store cell data.
 * @author dennis
 *
 */
public interface NDataGrid {

	public Object getValue(NCell cell);
	
	public void setValue(NCell cell, Object value);
	
	public boolean supportInsertDelete();
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int rowIdx, int size);
	public void deleteColumn(int rowIdx, int size);
	
	//TODO
	public boolean valideValue(NCell cell, Object value);
}
