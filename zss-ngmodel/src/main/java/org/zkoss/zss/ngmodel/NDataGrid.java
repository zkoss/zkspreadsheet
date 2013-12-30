package org.zkoss.zss.ngmodel;

import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCell.CellType;

/**
 * The data grid to store cell value when cell type is {@link CellType#BLANK},{@link CellType#BOOLEAN},{@link CellType#NUMBER} or
 * {@link CellType#STRING}. The runtime type {@link CellType#ERROR} and {@link CellType#FORMULA} are stored in {@link NCell} directly.
 * @author dennis
 *
 */
public interface NDataGrid {

	/**
	 * Get the cell value on row, column.
	 * @param row the row index
	 * @param column the column  index
	 * @return the cell value, excepted type is NULL, {@link CellType#BLANK},{@link CellType#BOOLEAN},{@link CellType#NUMBER} or
	 * {@link CellType#STRING}
	 */
	public NCellValue getValue(int row, int column);
	/**
	 * Set the cell value on row, column.
	 * @param row the row index
	 * @param column the column  index
	 * @param value  the cell value, excepted type is NULL (for clear), {@link CellType#BLANK},{@link CellType#BOOLEAN},{@link CellType#NUMBER} or
	 * {@link CellType#STRING}
	 */
	public void setValue(int row, int column, NCellValue value);
	
	/**
	 * Validate the cell value on row, column for {@link #setValue(int, int, NCellValue)}.
	 * Note, the cell value with type {@link CellType#ERROR} or {@link CellType#FORMULA} will also pass to this validation method
	 * to let you allow or reject if user types an formula or error on the cell, however it will not call
	 * {@link #setValue(int, int, NCellValue)} to stored this kind of cell value, it store on cell directly.
	 * @param row the row index
	 * @param column the column  index
	 * @param value  the cell value, excepted type is NULL (for clear), {@link CellType#BLANK},{@link CellType#BOOLEAN},{@link CellType#NUMBER} or
	 * {@link CellType#STRING}, it is also possible to be {@link CellType#ERROR} or {@link CellType#FORMULA} 
	 */
	public boolean validateValue(int row, int column, NCellValue value);
	
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
	
	//support start end index
	public boolean isProvidedStartEndIndex();
	public int getStartRowIndex();
	public int getEndRowIndex();
	public int getStartCellIndex(int rowIdx);
	public int getEndCellIndex(int rowIdx);
}
