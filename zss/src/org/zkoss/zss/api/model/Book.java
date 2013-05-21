package org.zkoss.zss.api.model;


public interface Book {
	public enum BookType {
		EXCEL_2003, EXCEL_2007
	}

	public Object getNative();
	
	public String getBookName();

	public BookType getType();

	public int getSheetIndex(Sheet sheet);

	public int getNumberOfSheets();

	public Sheet getSheetAt(int index);

	public Sheet getSheet(String name);

}
