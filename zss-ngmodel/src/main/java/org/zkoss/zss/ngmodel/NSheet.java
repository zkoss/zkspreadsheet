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
	public Iterator<NRow> getDataGridJoinedRowIterator();
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
	public Iterator<NCell> getDataGridJoinedCellIterator(int row);
	
	
	public int getDefaultRowHeight();
	public int getDefaultColumnWidth();
	
	public void setDefaultRowHeight(int height);
	public void setDefaultColumnWidth(int width);
	
	
	public NRow getRow(int rowIdx);
	
	NColumnArray getColumnArray(int columnIdx);
	public NColumn getColumn(int columnIdx);
	
	public NCell getCell(int rowIdx, int columnIdx);
	
	public String getId();
	
	public NViewInfo getViewInfo();
	public NPrintInfo getPrintInfo();
	
	//editable
//	public void clearRow(int rowIdx, int rowIdx2);
//	public void clearColumn(int columnIdx,int columnIdx2);
	public void clearCell(int rowIdx, int columnIdx,int rowIdx2,int columnIdx2);
	
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int columnIdx, int size);
	public void deleteColumn(int columnIdx, int size);
	
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
	public void removeMergedRegion(CellRegion region);
	public void addMergedRegion(CellRegion region);
	public int getNumOfMergedRegion();
	public CellRegion getMergedRegion(int idx);
	
	/**
	 * Get the merged region that overlapped the region
	 * @return the region that contains overlapped or null if not found.
	 */
	public List<CellRegion> getOverlappedMergedRegions(CellRegion region);
	
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

	
	public boolean isProtected();
	
	public void setProtected(boolean protect);

	public boolean isAutoFilterMode();

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
	
}
