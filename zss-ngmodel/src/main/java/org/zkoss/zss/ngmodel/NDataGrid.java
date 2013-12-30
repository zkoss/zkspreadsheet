package org.zkoss.zss.ngmodel;

import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCell.CellType;

/**
 * The data grid to store cell value.
 * @author dennis
 *
 */
public interface NDataGrid {

	/**
	 * Get the cell value on row, column.
	 * @param rowIdx the row index
	 * @param columnIdx the column  index
	 * @return the cell value
	 */
	public NCellValue getValue(int rowIdx, int columnIdx);
	/**
	 * Set the cell value on row, column.
	 * @param rowIdx the row index
	 * @param columnIdx the column  index
	 * @param value  the cell value
	 */
	public void setValue(int rowIdx, int columnIdx, NCellValue value);
	
	/**
	 * Validate the cell value on row, column for {@link #setValue(int, int, NCellValue)}.
	 * 
	 * @param rowIdx the row index
	 * @param columnIdx the column  index
	 * @param value  the cell value 
	 */
	public boolean validateValue(int rowIdx, int columnIdx, NCellValue value);
	
	//support operations
	public boolean isSupportedOperations();
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int rowIdx, int size);
	public void deleteColumn(int rowIdx, int size);
	
	//support data iterator
	public boolean isProvidedIterator();
	public Iterator<NDataRow> getRowIterator();
	public Iterator<NDataCell> getCellIterator(int row);
	
	public int getStartRowIndex();
	public int getEndRowIndex();
	public int getStartCellIndex(int rowIdx);
	public int getEndCellIndex(int rowIdx);
}
