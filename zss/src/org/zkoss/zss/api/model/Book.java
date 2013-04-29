package org.zkoss.zss.api.model;

import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;

public interface Book {
	public enum BookType {
		EXCEL_2003, EXCEL_2007
	}
	
	public Object getNative();

	public String getBookName();

	public BookType getType();

	public int getSheetIndex(Sheet sheet);

	public Font findFont(Boldweight boldweight, Color color,
			short fontHeight, String fontName, boolean italic,
			boolean strikeout, TypeOffset typeOffset, Underline underline);

	public Color getColorFromHtmlColor(String htmlColor);

	public int getNumberOfSheets();

	public Sheet getSheetAt(int index);

	public Sheet getSheet(String name);

}
