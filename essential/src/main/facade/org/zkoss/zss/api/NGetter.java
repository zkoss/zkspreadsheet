package org.zkoss.zss.api;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NColor;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.TypeOffset;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Worksheet;

public class NGetter {

	Range range;

	public NGetter(Range range) {
		this.range = range;
	}

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, default cell style if cell not found
	 */
	public NCellStyle getCellStyle() {
		Worksheet sheet = range.getSheet();
		int r = range.getRow();
		int c = range.getColumn();
		Row row = sheet.getRow(r);
		if (row == null)
			return getDefaultCellStyle();
		Cell cell = row.getCell(c);
		if (cell == null)
			return getDefaultCellStyle();
		return new NCellStyle(sheet.getBook(), cell.getCellStyle());
	}

	/**
	 * get default cell style of book of this range.
	 * 
	 * @return
	 */
	public NCellStyle getDefaultCellStyle() {
		Book book = range.getSheet().getBook();
		return new NCellStyle(book, book.getCellStyleAt((short) 0));
	}

	public NFont findFont(Boldweight boldweight, NColor color, short fontHeight,
			String fontName, boolean italic, boolean strikeout,
			TypeOffset typeOffset, Underline underline) {
		Book book = range.getSheet().getBook();
		Font font = book.findFont(EnumUtil.toFontBoldweight(boldweight), color.getNative(), fontHeight, fontName,
				italic, strikeout, EnumUtil.toFontTypeOffset(typeOffset), EnumUtil.toFontUnderline(underline));
		return font==null?null:new NFont(book,font);
	}

}
