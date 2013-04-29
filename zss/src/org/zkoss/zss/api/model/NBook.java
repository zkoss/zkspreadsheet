package org.zkoss.zss.api.model;

import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.TypeOffset;
import org.zkoss.zss.api.model.NFont.Underline;

public interface NBook {
	public enum BookType {
		EXCEL_2003, EXCEL_2007
	}
	
	public Object getNative();

	public String getBookName();

	public BookType getType();

	public int getSheetIndex(NSheet sheet);

	public NFont findFont(Boldweight boldweight, NColor color,
			short fontHeight, String fontName, boolean italic,
			boolean strikeout, TypeOffset typeOffset, Underline underline);

	public NColor getColorFromHtmlColor(String htmlColor);

	public int getNumberOfSheets();

	public NSheet getSheetAt(int index);

	public NSheet getSheet(String name);

}
