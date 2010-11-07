package org.zkoss.zss.app.cell;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;

public class CellHelper {
	private CellHelper(){};
	
	/**
	 * Returns cell font family
	 * @param cell
	 * @return
	 */
	public static Font getFont(Cell cell) {
		if (cell == null)
			return null;
		
		Book book = (Book)cell.getSheet().getWorkbook();
		CellStyle cellStyle = cell.getCellStyle();
		return book.getFontAt(cellStyle.getFontIndex());
	}
	
	public static boolean isBold(Font font) {
		return font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
	}
	
	public static boolean isAlignLeft(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return cellStyle.getAlignment() == CellStyle.ALIGN_LEFT;
	}
	
	public static boolean isAlignCenter(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return cellStyle.getAlignment() == CellStyle.ALIGN_CENTER;
	}

	public static boolean isAlignRight(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return cellStyle.getAlignment() == CellStyle.ALIGN_RIGHT;
	}
	
	public static String getFontHTMLColor(Cell cell, Font font) {
		Book book = (Book)cell.getSheet().getWorkbook();
		String color = BookHelper.getFontHTMLColor(book, font);
		if (color == null || BookHelper.AUTO_COLOR.equals(color))
			return "#000000";
		return  color;
	}
	
	public static String getBackgroundHTMLColor(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		Book book = (Book)cell.getSheet().getWorkbook();
		String color = BookHelper.colorToHTML(book, cellStyle.getFillForegroundColorColor());
		if (color == null || BookHelper.AUTO_COLOR.equals(color))
			return "#FFFFFF";
		return color;
	}
}
