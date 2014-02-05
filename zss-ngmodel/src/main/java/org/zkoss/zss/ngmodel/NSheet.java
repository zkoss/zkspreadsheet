/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author dennis
 * @since 3.5.0
 */
public interface NSheet {
	
	/**
	 * Get the owner book
	 * @return the owner book
	 */
	public NBook getBook();
	
	/**
	 * Get the sheet name
	 * @return the sheet name
	 */
	public String getSheetName();
	
	public Iterator<NRow> getRowIterator();
	public Iterator<NRow> getRowIterator(boolean joinDataGrid);
	public Iterator<NColumn> getColumnIterator();
	public Iterator<NColumnArray> getColumnArrayIterator();
	/**
	 * Set up a column array, if any array are overlapped, it throws IllegalStateException.
	 * If you setup a column array that is not continue, (for example, 0-2, 5-6), then it will create a continue array automatically.(3-4 in the example) 
	 * @param colunmIdx 
	 * @param lastColumnIdx
	 * @return the new created column array
	 */
	public NColumnArray setupColumnArray(int colunmIdx,int lastColumnIdx);	
	public Iterator<NCell> getCellIterator(int row);
	public Iterator<NCell> getCellIterator(int row,boolean joinDataGrid);
	
	
	public int getDefaultRowHeight();
	public int getDefaultColumnWidth();
	
	public void setDefaultRowHeight(int height);
	public void setDefaultColumnWidth(int width);
	
	
	public NRow getRow(int rowIdx);
	
	NColumnArray getColumnArray(int columnIdx);
	public NColumn getColumn(int columnIdx);
	
	public NCell getCell(int rowIdx, int columnIdx);
	public NCell getCell(String cellRefString);
	
	public String getId();
	
	public NSheetViewInfo getViewInfo();
	public NPrintSetup getPrintSetup();
	
	public abstract int getStartRowIndex();
	public abstract int getEndRowIndex();
	public abstract int getStartColumnIndex();
	public abstract int getEndColumnIndex();
	public abstract int getStartCellIndex(int rowIdx);
	public abstract int getEndCellIndex(int rowIdx);
	
	
	public abstract int getStartRowIndex(boolean joinDataGrid);
	public abstract int getEndRowIndex(boolean joinDataGrid);
	public abstract int getStartCellIndex(int rowIdx,boolean joinDataGrid);
	public abstract int getEndCellIndex(int rowIdx,boolean joinDataGrid);
	//editable
//	public void clearRow(int rowIdx, int rowIdx2);
//	public void clearColumn(int columnIdx,int columnIdx2);
	public void clearCell(int rowIdx, int columnIdx,int lastRowIdx,int lastColumnIdx);
	public void clearCell(CellRegion region);
	
	public void moveCell(int rowIdx, int columnIdx,int lastRowIdx,int lastColumnIdx, int rowOffset, int columnOffset);
	public void moveCell(CellRegion region, int rowOffset, int columnOffset);
	
	public void insertRow(int rowIdx, int lastRowIdx);
	public void deleteRow(int rowIdx, int lastRowIdx);
	public void insertColumn(int columnIdx, int lastColumnIdx);
	public void deleteColumn(int columnIdx, int lastColumnIdx);
	public void insertCell(int rowIdx,int columnIdx,int lastRowIndex, int lastColumnIndex,boolean horizontal);
	public void insertCell(CellRegion region,boolean horizontal);
	public void deleteCell(CellRegion region,boolean horizontal);
	public void deleteCell(int rowIdx,int columnIdx,int lastRowIndex, int lastColumnIndex,boolean horizontal);
	
	public NPicture addPicture(NPicture.Format format, byte[] data, NViewAnchor anchor);
	public NPicture getPicture(String picid);
	public void deletePicture(NPicture picture);
	public int getNumOfPicture();
	public NPicture getPicture(int idx);
	public List<NPicture> getPictures();
	
	
	public NChart addChart(NChart.NChartType type, NViewAnchor anchor);
	public NChart getChart(String chartid);
	public void deleteChart(NChart chart);
	public int getNumOfChart();
	public NChart getChart(int idx);
	public List<NChart> getCharts();
	
	
	public List<CellRegion> getMergedRegions();
	/**
	 * Remove the merged area that are contained by region 
	 * @param region
	 * @param removeOverlpas true if you want to remove the merged areas that are just overlapped.
	 */
	public void removeMergedRegion(CellRegion region,boolean removeOverlaps);
	/**
	 * Add a merged area, you can't assign a area that overlaps existed merged area.
	 * @param region
	 */
	public void addMergedRegion(CellRegion region);
	public int getNumOfMergedRegion();
	public CellRegion getMergedRegion(int idx);
	
	/**
	 * Get the merged region that overlapped the region
	 * @return the regions that overlaps
	 */
	public List<CellRegion> getOverlapsMergedRegions(CellRegion region, boolean excludeContains);
	/**
	 * Get the merged region that are contained by region.
	 * @return the regions that are contained
	 */
	public List<CellRegion> getContainsMergedRegions(CellRegion region);
	public CellRegion getMergedRegion(int row,int column);
	public CellRegion getMergedRegion(String cellRefString);
	
	public NDataValidation addDataValidation(CellRegion region);
	public NDataValidation getDataValidation(String id);
	public void deleteDataValidation(NDataValidation validation);
	public int getNumOfDataValidation();
	public NDataValidation getDataValidation(int idx);
	public List<NDataValidation> getDataValidations();
	
	/**
	 * @param row
	 * @param column
	 * @return the first data validation at row, column
	 */
	public NDataValidation getDataValidation(int row, int column);
	
	/**
	 * Get the runtime custom attribute that stored in this sheet
	 * @param name the attribute name
	 * @return the value, or null if not found
	 */
	public Object getAttribute(String name);
	
	/**
	 * Set the runtime custom attribute to stored in this sheet, the attribute is only use for developer to stored runtime data in the sheet,
	 * values will not stored to excel when exporting.
	 * @param name name the attribute name
	 * @param value the attribute value
	 */
	public Object setAttribute(String name,Object value);
	
	/**
	 * Get the unmodifiable runtime attributes map
	 * @return
	 */
	public Map<String,Object> getAttributes();

	/**
	 * Check if the sheet is protected
	 * @return
	 */
	public boolean isProtected();
	
	/**
	 * Sets password to protect sheet, set null to unprotect it.
	 * @param protection
	 */
	public void setPassword(String password);
	
	/**
	 * Gets the password
	 * @return
	 */
	public String getPassword();

	/**
	 * Get the data grid that this sheet stores it data, by default it is null, and the data store on cell directly
	 * @return
	 */
	public NDataGrid getDataGrid();
	
	/**
	 * Sets the data grid to store data of this sheet
	 * @param dataGrid
	 */
	public void setDataGrid(NDataGrid dataGrid);
	
	
	/**
	 * Gets the auto filter information if there is.
	 * @return the auto filter, or null if not found
	 */
	public NAutoFilter getAutoFilter();
	
	/**
	 * Creates a new auto filter, the old one will be drop directly.
	 * @param region the auto filter region
	 * @return the new auto filter.
	 */
	public NAutoFilter createAutoFilter(CellRegion region);
	
	
	/**
	 * Delete current autofilter if it has
	 */
	public void deleteAutoFilter();
	
	/**
	 * Clear auto filter if there is.
	 */
	public void clearAutoFilter();
	
	/**
	 * paste cell from src sheet to this sheet, the sheets must in same book
	 * @param src src sheet and it's region to paste
	 * @param dest destination region in this sheet
	 * @param option the copy option
	 * @return the final effected region
	 */
	public CellRegion pasteCell(SheetRegion src,CellRegion dest,PasteOption option);
}
