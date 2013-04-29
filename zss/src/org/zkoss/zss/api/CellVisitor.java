package org.zkoss.zss.api;


public interface CellVisitor {

	/**
	 * Should ignore to visit a cell on (row,column) if the cell doesn't existed.
	 * This method is called only when the cell of (row,column) is null.
	 * If return false, caller should create cell before call {@code #visit(Range)} 
	 * @param row the row of the cell
	 * @param column the column of the cell
	 * @return true if ignore to visit the cell if it is not exist.
	 */
	public boolean ignoreIfNotExist(int row, int column);

	/**
	 * Visits the cell
	 * @param cellRange the range of a cell
	 * @return true if should continue visit next cell.
	 */
	public boolean visit(Range cellRange);
}
