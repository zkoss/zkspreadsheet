/* StyleUtil.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jun 16, 2008 2:50:27 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.util.logging.Log;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
/**
 * A utility class to help spreadsheet set style of a cell
 * @author Dennis.Chen
 *
 */
public class Styles {
	private static final Log log = Log.lookup(Styles.class);
	
	public static CellStyle cloneCellStyle(Cell cell) {
		final CellStyle destination = cell.getSheet().getWorkbook().createCellStyle();
		destination.cloneStyleFrom(cell.getCellStyle());
		return destination;
	}
	
	public static void setFontColor(Worksheet sheet, int row, int col, String color){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final Color orgColor = BookHelper.getFontColor(book, font);
		final Color newColor = BookHelper.HTMLToColor(book, color);
		if (orgColor == newColor || orgColor != null && orgColor.equals(newColor)) {
			return;
		}
		final short boldWeight = font.getBoldweight();
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, newColor, fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setFillColor(Worksheet sheet, int row, int col, String color){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final Color orgColor = cell.getCellStyle().getFillForegroundColorColor();
		final Color newColor = BookHelper.HTMLToColor(book, color);
		if (orgColor == newColor || orgColor != null  && orgColor.equals(newColor)) { //no change, skip
			return;
		}
		final CellStyle style = cloneCellStyle(cell);
		BookHelper.setFillForegroundColor(style, newColor);
		cell.setCellStyle(style);
	}
	
	public static void setTextWrap(Worksheet sheet,int row,int col,boolean wrap){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final boolean textWrap = cell.getCellStyle().getWrapText();
		if (wrap == textWrap) { //no change, skip
			return;
		}
		final CellStyle style = cloneCellStyle(cell);
		style.setWrapText(wrap);
		cell.setCellStyle(style);
	}
	
	public static void setFontSize(Worksheet sheet,int row,int col,int fontHeight){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final short orgSize = font.getFontHeight();
		if (orgSize == fontHeight) { //no change, skip
			return;
		}
		final short boldWeight = font.getBoldweight();
		final Color color = BookHelper.getFontColor(book, font);
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, color, (short)fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setFontStrikethrough(Worksheet sheet,int row,int col, boolean strikeout){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final boolean orgStrikeout = font.getStrikeout();
		if (orgStrikeout == strikeout) { //no change, skip
			return;
		}
		final short boldWeight = font.getBoldweight();
		final Color color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setFontType(Worksheet sheet,int row,int col,String name){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final String orgName = font.getFontName();
		if (orgName.equals(name)) { //no change, skip
			return;
		}
		final short boldWeight = font.getBoldweight();
		final Color color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setBorder(Worksheet sheet,int row,int col, String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0xF);
	}
	public static void setBorderTop(Worksheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x4);
	}
	public static void setBorderLeft(Worksheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x8);
	}
	public static void setBorderBottom(Worksheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x1);
	}
	public static void setBorderRight(Worksheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x2);
	}
	public static void setBorder(Worksheet sheet,int row,int col, short color, short lineStyle, int at){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final CellStyle style = cloneCellStyle(cell);
		if((at & BookHelper.BORDER_EDGE_LEFT)!=0) {
			style.setBorderLeft(lineStyle);
		}
		if((at & BookHelper.BORDER_EDGE_TOP)!=0){
			style.setTopBorderColor(color);
			style.setBorderTop(lineStyle);
		}
		if((at & BookHelper.BORDER_EDGE_RIGHT)!=0){
			style.setRightBorderColor(color);
			style.setBorderRight(lineStyle);
		}
		if((at & BookHelper.BORDER_EDGE_BOTTOM)!=0){
			style.setBottomBorderColor(color);
			style.setBorderBottom(lineStyle);
		}
		cell.setCellStyle(style);
	}
	
	public static void setBorder(Worksheet sheet,int row,int col, Color color, short lineStyle, int at){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final CellStyle style = cloneCellStyle(cell);
		if((at & BookHelper.BORDER_EDGE_LEFT)!=0) {
			BookHelper.setLeftBorderColor(style, color);
			style.setBorderLeft(lineStyle);
		}
		if((at & BookHelper.BORDER_EDGE_TOP)!=0){
			BookHelper.setTopBorderColor(style, color);
			style.setBorderTop(lineStyle);
		}
		if((at & BookHelper.BORDER_EDGE_RIGHT)!=0){
			BookHelper.setRightBorderColor(style, color);
			style.setBorderRight(lineStyle);
		}
		if((at & BookHelper.BORDER_EDGE_BOTTOM)!=0){
			BookHelper.setBottomBorderColor(style, color);
			style.setBorderBottom(lineStyle);
		}
		cell.setCellStyle(style);
	}
	
	public static void setFontBoldWeight(Worksheet sheet,int row,int col,short boldWeight){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final short orgBoldWeight = font.getBoldweight();
		if (orgBoldWeight == boldWeight) { //no change, skip
			return;
		}
		final Color color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setFontItalic(Worksheet sheet, int row, int col, boolean italic) {
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final boolean orgItalic = font.getItalic();
		if (orgItalic == italic) { //no change, skip
			return;
		}
		final short boldWeight = font.getBoldweight();
		final Color color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setFontUnderline(Worksheet sheet,int row,int col, byte underline){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final Book book = (Book) sheet.getWorkbook();
		final short fontIdx = cell.getCellStyle().getFontIndex();
		final Font font = book.getFontAt(fontIdx);
		final byte orgUnderline = font.getUnderline();
		if (orgUnderline == underline) { //no change, skip
			return;
		}
		final short boldWeight = font.getBoldweight();
		final Color color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final Font newFont = BookHelper.getOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		final CellStyle style = cloneCellStyle(cell);
		style.setFont(newFont);
		cell.setCellStyle(style);
	}
	
	public static void setTextHAlign(Worksheet sheet,int row,int col, short align){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final short orgAlign = cell.getCellStyle().getAlignment();
		if (align == orgAlign) { //no change, skip
			return;
		}
		final CellStyle style = cloneCellStyle(cell);
		style.setAlignment(align);
		cell.setCellStyle(style);
	}
	
	public static void setTextVAlign(Worksheet sheet,int row,int col, short valign){
		final Cell cell = Utils.getOrCreateCell(sheet,row,col);
		final short orgValign = cell.getCellStyle().getVerticalAlignment();
		if (valign == orgValign) { //no change, skip
			return;
		}
		final CellStyle style = cloneCellStyle(cell);
		style.setAlignment(valign);
		cell.setCellStyle(style);
	}
	
}
