package org.zkoss.zss.api.model;

import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;

public interface Book {
	public enum BookType {
		EXCEL_2003, EXCEL_2007
	}

	public String getBookName();

	public BookType getType();

	public int getSheetIndex(Sheet sheet);

	public int getNumberOfSheets();

	public Sheet getSheetAt(int index);

	public Sheet getSheet(String name);

}
