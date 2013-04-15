package org.zkoss.zss.api;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Color;
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
import org.zkoss.zss.model.impl.BookHelper;

public class NGetter {

	Range range;

	public NGetter(Range range) {
		this.range = range;
	}

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, the check row style and column cell style if cell not found, if row and column style is not exist, then return default style of sheet
	 */
	public NCellStyle getCellStyle() {
		Worksheet sheet = range.getSheet();
		Book book = sheet.getBook();
		
		int r = range.getRow();
		int c = range.getColumn();
		CellStyle style = null;
		Row row = sheet.getRow(r);
		if (row != null){
			Cell cell = row.getCell(c);
			
			if (cell != null){//cell style
				style = cell.getCellStyle();
			}
			if(style==null && row.isFormatted()){//row sytle
				style = row.getRowStyle();
			}
		}
		if(style==null){//col style
			style = sheet.getColumnStyle(c);
		}
		if(style==null){//default
			style = book.getCellStyleAt((short) 0);
		}
		
		return new NCellStyle(sheet.getBook(), style);		
	}

//	/**
//	 * get default cell style of book of this range.
//	 * 
//	 * @return
//	 */
//	public NCellStyle getDefaultCellStyle() {
//		Book book = range.getSheet().getBook();
//		return new NCellStyle(book, book.getCellStyleAt((short) 0));
//	}

	public NFont findFont(Boldweight boldweight, NColor color, short fontHeight,
			String fontName, boolean italic, boolean strikeout,
			TypeOffset typeOffset, Underline underline) {
		Book book = range.getSheet().getBook();
		Font font;
		
		font = book.findFont(EnumUtil.toFontBoldweight(boldweight), color.getNative(), fontHeight, fontName,
				italic, strikeout, EnumUtil.toFontTypeOffset(typeOffset), EnumUtil.toFontUnderline(underline));
//		font = BookHelper.getOrCreateFont(book,EnumUtil.toFontBoldweight(boldweight), color.getNative(), fontHeight, fontName,
//				italic, strikeout, EnumUtil.toFontTypeOffset(typeOffset), EnumUtil.toFontUnderline(underline));
		return font==null?null:new NFont(book,font);
	}

	public NColor getColorFromHtmlColor(String htmlColor) {
		Book book = range.getSheet().getBook();
		Color color = BookHelper.HTMLToColor(book, htmlColor);//never null
		return new NColor(book,color);
	}

}
