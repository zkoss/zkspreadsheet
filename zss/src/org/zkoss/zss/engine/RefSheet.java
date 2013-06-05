/* RefSheet.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 2:52:40 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine;

import java.util.Set;

import org.zkoss.zss.engine.impl.CellRefImpl;

/**
 * Internal User Only. Reference areas(see {@link Ref}) in a sheet for hit test. When user change value of a cell,
 * the engine shall find which {@link Ref}s includes this cell(a hit).
 * 
 * @author henrichen
 */
public interface RefSheet {
	/**
	 * Returns the sheet name.
	 * @return the sheet name.
	 */
	public String getSheetName();
	/**
	 * Returns the owner reference book of this reference sheet.
	 * @return the owner reference book of this reference sheet.
	 */
	public RefBook getOwnerBook();
	/**
	 * Given a point, return the hit {@link Ref} set.
	 * @param row the row index of the cell
	 * @param col the column index of the cell
	 * @return the hit {@link Ref} set.
	 */
	public Set<Ref> getHitRefs(int row, int col);
	
	/**
	 * Returns the reference area from the reference sheet; return null if not exists.
	 * @param tRow top row index   
	 * @param lCol left column index   
	 * @param bRow bottom row index   
	 * @param rCol right column index   
	 * @return a {@link Ref} from the reference sheet; return null if not exists. 
	 */
	public Ref getRef(int tRow, int lCol, int bRow, int rCol);
	 
	/**
	 * Returns or create a reference area(if not exist) from the reference sheet. 
	 * @param tRow top row index   
	 * @param lCol left column index   
	 * @param bRow bottom row index   
	 * @param rCol right column index   
	 * @return a {@link Ref} from the reference sheet; create one if not exists. 
	 */
	public Ref getOrCreateRef(int tRow, int lCol, int bRow, int rCol);
	
	/**
	 * Remove and return a reference area from the reference sheet. If not exists, return null; 
	 * @param tRow top row index   
	 * @param lCol left column index   
	 * @param bRow bottom row index   
	 * @param rCol right column index   
	 * @return the removed {@link Ref}; or null if no such reference
	 */
	public Ref removeRef(int tRow, int lCol, int bRow, int rCol);
	
	/**
	 * Create source cell reference to precedent reference area dependency.
	 * @param srcRow source cell reference row
	 * @param srcCol Source cell reference column
	 * @param sheet precedent reference sheet
	 * @param tRow precedent reference area top row
	 * @param lCol precedent reference area left column
	 * @param bRow precedent reference area bottom row
	 * @param rCol precedent reference area right column
	 */
	public void addDependency(int srcRow, int srcCol, RefSheet sheet, int tRow, int lCol, int bRow, int rCol);
	
	/**
	 * Remove source cell reference to precedent reference area dependency.
	 * @param srcRow source cell reference row
	 * @param srcCol Source cell reference column
	 * @param sheet precedent reference sheet
	 * @param tRow precedent reference area top row
	 * @param lCol precedent reference area left column
	 * @param bRow precedent reference area bottom row
	 * @param rCol precedent reference area right column
	 */
	public void removeDependency(int srcRow, int srcCol, RefSheet sheet, int tRow, int lCol, int bRow, int rCol);

	/**
	 * Create source cell reference to precedent "variable name" dependency.
	 * @param srcRow source cell reference row
	 * @param srcCol Source cell reference column
	 * @param name the variable name that was the precedent of this cell
	 */
	public void addDependency(int srcRow, int srcCol, String name);

	/**
	 * Remove source cell reference to precedent "variable name" dependency. 
	 * @param srcRow source cell reference row
	 * @param srcCol Source cell reference column
	 * @param name the variable name that was the precedent of this cell
	 */
	public void removeDependency(int srcRow, int srcCol, String name);
	
	/**
	 * Returns the direct dependent cell references that are affected by the the cell 
	 * at the specified row and column.
	 * @param row row index
	 * @param col column index
	 * @return the direct dependent cell references that is affected by the cell 
	 * at the specified row and column.  
	 */
	public Set<Ref> getDirectDependents(int row, int col);

	/** 
	 * Returns all dependent cell references that are affected by the the cell at the 
	 * specified row and column.
	 * @param row row index
	 * @param col column index
	 * @return the direct and indirect dependent cell references that is affected by  
	 * the cell at the specified row and column.  
	 */
	public Set<Ref> getAllDependents(int row, int col);
	
	/** 
	 * Returns last dependent cell references that are affected by the the cell at 
	 * the specified row and column.
	 * @param row row index
	 * @param col column index
	 * @return the last dependent cell references that is affected by the cell 
	 * at the specified row and column.  
	 */
	public Set<Ref> getLastDependents(int row, int col);
	
	/**
	 * Returns both all and the last dependent cell references that are affected by
	 * the cell at the specified row and column.
	 * @param row row index
	 * @param col column index
	 * @return reference set: [0] the "last" dependent cell references to be re-evaluated the associated cell value; 
	 * 	[1] the "all" dependent cell references to be reloaded the associated cell value.
	 */
	public Set<Ref>[] getBothDependents(int row, int col);
	
	/** 
	 * Returns all direct precedent cell references that affects the cell at the specified row and column.
	 * @param row row index
	 * @param col column index
	 * @return the direct precedent cell references affects the cell at the specified row and column.  
	 */
	public Set<Ref> getDirectPrecedents(int row, int col);
	
	/** 
	 * Returns all precedent cell references that affects the cell at the 
	 * specified row and column.
	 * @param row row index
	 * @param col column index
	 * @return the direct and indirect precedent cell references affects the cell at the specified row and column.  
	 */
	public Set<Ref> getAllPrecedents(int row, int col);
	
	/**
	 * Insert number of rows from the specified row index. 
	 * @param row the insertion point of row 
	 * @param num the number of rows to insert
	 * @return the last affected references(for re-evaluation, at [0]) and all affected references(for re-render, at [1]) 
	 */
	public Set<Ref>[] insertRows(int row, int num);
	
	/**
	 * Delete number of rows from the specified row index. 
	 * @param row the removing point of row 
	 * @param num the number of rows to remove
	 * @return the last affected references(for re-evaluation, at [0]) and all affected references(for re-render, at [1]) 
	 */
	public Set<Ref>[] deleteRows(int row, int num);
	
	/**
	 * Insert number of columns from the specified column index. 
	 * @param col the insertion point of column
	 * @param num the number of columns to insert
	 * @return the last affected references(for re-evaluation, at [0]) and all affected references(for re-render, at [1]) 
	 */
	public Set<Ref>[] insertColumns(int col, int num);

	/**
	 * Delete number of columns from the specified row index. 
	 * @param col the removing point of column 
	 * @param num the number of rows to remove
	 * @return the last affected references(for re-evaluation, at [0]) and all affected references(for re-render, at [1]) 
	 */
	public Set<Ref>[] deleteColumns(int col, int num);
	
	/**
	 * Insert a range of cells.
	 * @param tRow top row index of the range
	 * @param lCol left column index of the range
	 * @param bRow bottom row index of the range
	 * @param rCol right column index of the range
	 * @param horizontal neighbor cells shall move to right(true)/bottom(false)
	 * @return the last affected references(for re-evaluation, at [0]) and all affected references(for re-render, at [1]) 
	 */
	public Set<Ref>[] insertRange(int tRow, int lCol, int bRow, int rCol, boolean horizontal);
	
	/**
	 * Delete a range of cells.
	 * @param tRow top row index of the range
	 * @param lCol left column index of the range
	 * @param bRow bottom row index of the range
	 * @param rCol right column index of the range
	 * @param horizontal neighbor cells shall move to left(true)/top(false)
	 * @return the last affected references(for re-evaluation, at [0]) and all affected references(for re-render, at [1]) 
	 */
	public Set<Ref>[] deleteRange(int tRow, int lCol, int bRow, int rCol, boolean horizontal);
	
	/**
	 * Move a range of cells to a new position as specified by nCol(positive to move right, negative to move left) 
	 * and nRow(positive to move down, negative to move up)
	 * 
	 * @param tRow top row index of the range
	 * @param lCol left column index of the range
	 * @param bRow bottom row index of the range
	 * @param rCol right column index of the range
	 * @param nRow number of rows to move
	 * @param nCol number of columns to move
	 * @return the to be evaluated references(for re-evaluation, at [0]) and all affected references(for re-render, at [1])
	 */
	public Set<Ref>[] moveRange(int tRow, int lCol, int bRow, int rCol, int nRow, int nCol);

	/**
	 * Sets whether the given reference contains indirect precedents.
	 * @param row top row index of the range
	 * @param col left column index of the range
	 * @param withIndirectPrecedent whether the reference contains indirect precedents.
	 */
	public void setRefWithIndirectPrecedent(int row, int col, boolean withIndirectPrecedent);
	
	
	public Set<Ref> removeExternalRef();
}
