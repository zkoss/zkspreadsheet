package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Workbook;


public interface Book {
	public enum BookType {
		EXCEL_2003, EXCEL_2007
	}

	public Workbook getPoiBook();
	
	public String getBookName();

	public BookType getType();

	public int getSheetIndex(Sheet sheet);

	public int getNumberOfSheets();

	public Sheet getSheetAt(int index);

	public Sheet getSheet(String name);

	public void setShareScope(String scope);
	
	public String getShareScope();

}
