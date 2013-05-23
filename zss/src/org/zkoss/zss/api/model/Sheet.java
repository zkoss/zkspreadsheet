package org.zkoss.zss.api.model;

import java.util.List;


public interface Sheet {

	public org.zkoss.poi.ss.usermodel.Sheet getPoiSheet();
	
	public Book getBook();

	public boolean isProtected();

	public boolean isAutoFilterEnabled();

	public boolean isDisplayGridlines();
	
	public boolean isRowHidden(int row);
	
	public boolean isColumnHidden(int column);

	public String getSheetName();
	
	public List<Chart> getCharts();
	
	public List<Picture> getPictures();
	
	public int getRowFreeze();
	
	public int getColumnFreeze();

	public boolean isPrintGridlines();
	
	/**
	 * get row height in pixel
	 * @param row
	 * @return row height in pixel
	 */
	public int getRowHeight(int row);
	
	/**
	 * get column width in pixel
	 * @param column
	 * @return column width in pixel
	 */
	public int getColumnWidth(int column);
	

}
