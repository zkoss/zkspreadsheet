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
package org.zkoss.zss.ngapi.impl;

import java.util.HashMap;

import org.zkoss.util.logging.Log;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.CellStyleMatcher;
import org.zkoss.zss.ngmodel.util.FontMatcher;
/**
 * A utility class to help spreadsheet set style of a cell
 * @author Dennis.Chen
 *
 */
public class NStyles {
//	private static final Log log = Log.lookup(NStyles.class);
	
	public static NCellStyle cloneCellStyle(NCell cell) {
		final NCellStyle destination = cell.getSheet().getBook().createCellStyle(cell.getCellStyle(), true);
		return destination;
	}
	
	public static void setFontColor(NSheet sheet, int row, int col, String color/*,HashMap<Integer,NCellStyle> cache*/){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont font = orgStyle.getFont();
		final NColor orgColor = font.getColor();
		final NColor newColor = book.createColor(color);
		if (orgColor == newColor || orgColor != null && orgColor.equals(newColor)) {
			return;
		}
		
//		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
//		if(hitStyle!=null){
//			cell.setCellStyle(hitStyle);
//			return;
//		}
		
		FontMatcher fontmatcher = new FontMatcher(font);
		fontmatcher.setColor(color);
		
		font = book.searchFont(fontmatcher);
		
		
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(font,true);
			font.setColor(book.createColor(color));
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
		
//		if(cache!=null){
//			cache.put((int)orgStyle.getIndex(), style);
//		}
	}
	
	/*
	public static void setFillColor(NSheet sheet, int row, int col, String htmlColor,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final NColor orgColor = orgStyle.getFillForegroundColorColor();
		final NColor newColor = BookHelper.HTMLToColor(book, htmlColor);
		if (orgColor == newColor || orgColor != null  && orgColor.equals(newColor)) { //no change, skip
			return;
		}
		
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		
		CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
		matcher.setFillForegroundColor(htmlColor);
		matcher.setFillPattern(CellStyle.SOLID_FOREGROUND);
		NCellStyle style = findStyle(sheet.getBook(), matcher);
		if(style==null){
			style  = cloneCellStyle(cell);
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			BookHelper.setFillForegroundColor(style, newColor);
		}
		cell.setCellStyle(style);
		
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
		
	}
	
	public static void setTextWrap(NSheet sheet,int row,int col,boolean wrap,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final boolean textWrap = orgStyle.getWrapText();
		if (wrap == textWrap) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		
		CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
		matcher.setWrapText(wrap);
		NCellStyle style = findStyle(sheet.getBook(), matcher);
		if(style==null){
			style  = cloneCellStyle(cell);
			style.setWrapText(wrap);
		}
		cell.setCellStyle(style);
		
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setFontHeight(NSheet sheet,int row,int col,int fontHeight,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short fontIdx = orgStyle.getFontIndex();
		final NFont font = book.getFontAt(fontIdx);
		final short orgSize = font.getFontHeight();
		if (orgSize == fontHeight) { //no change, skip
			return;
		}
		
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		
		final short boldWeight = font.getBoldweight();
		final NColor color = BookHelper.getFontColor(book, font);
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Object[] newFont = BookHelper.findOrCreateFont(book, boldWeight, color, (short)fontHeight, name, italic, strikeout, typeOffset, underline);
		
		NCellStyle style = null;
		if(Boolean.FALSE.equals(newFont[1])){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
			matcher.setFontIndex(((Font)newFont[0]).getIndex());
			style = findStyle(sheet.getBook(), matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(((Font)newFont[0]));
		}
		cell.setCellStyle(style);
		
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setFontStrikethrough(NSheet sheet,int row,int col, boolean strikeout,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short fontIdx = orgStyle.getFontIndex();
		final NFont font = book.getFontAt(fontIdx);
		final boolean orgStrikeout = font.getStrikeout();
		if (orgStrikeout == strikeout) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		final short boldWeight = font.getBoldweight();
		final NColor color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Object[] newFont = BookHelper.findOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		
		NCellStyle style = null;
		if(Boolean.FALSE.equals(newFont[1])){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
			matcher.setFontIndex(((Font)newFont[0]).getIndex());
			style = findStyle(sheet.getBook(), matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(((Font)newFont[0]));
		}
		cell.setCellStyle(style);
		
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setFontName(NSheet sheet,int row,int col,String name,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short fontIdx = orgStyle.getFontIndex();
		final NFont font = book.getFontAt(fontIdx);
		final String orgName = font.getFontName();
		if (orgName.equals(name)) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		
		final short boldWeight = font.getBoldweight();
		final NColor color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Object[] newFont = BookHelper.findOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);

		NCellStyle style = null;
		if(Boolean.FALSE.equals(newFont[1])){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
			matcher.setFontIndex(((Font)newFont[0]).getIndex());
			style = findStyle(sheet.getBook(), matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(((Font)newFont[0]));
		}
		cell.setCellStyle(style);
		
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setBorder(NSheet sheet,int row,int col, String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0xF);
	}
	public static void setBorderTop(NSheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x4);
	}
	public static void setBorderLeft(NSheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x8);
	}
	public static void setBorderBottom(NSheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x1);
	}
	public static void setBorderRight(NSheet sheet,int row,int col,String color, short linestyle){
		setBorder(sheet,row,col, BookHelper.HTMLToColor(sheet.getWorkbook(), color), linestyle, 0x2);
	}
	
	@Deprecated
	public static void setBorder(NSheet sheet,int row,int col, short color, short lineStyle, int at){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final NCellStyle style = cloneCellStyle(cell);
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
	
	public static void setBorder(NSheet sheet,int row,int col, NColor color, short lineStyle, int at){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = sheet.getBook();
		//ZSS-464 try to search existed matched style
		String colorHtml = BookHelper.colorToBorderHTML(book, color);
		NCellStyle style = null;
		boolean hasBorder = lineStyle != CellStyle.BORDER_NONE;
		if(colorHtml!=null){
			final NCellStyle oldstyle = NStyles.getCellStyle(cell);
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),oldstyle);
			if((at & BookHelper.BORDER_EDGE_LEFT)!=0) {
				if(hasBorder)
					matcher.setLeftBorderColor(colorHtml);
				else
					matcher.removeLeftBorderColor();
				
				matcher.setBorderLeft(lineStyle);
			}
			if((at & BookHelper.BORDER_EDGE_TOP)!=0){
				if(hasBorder) 
					matcher.setTopBorderColor(colorHtml);
				else
					matcher.removeTopBorderColor();
				
				matcher.setBorderTop(lineStyle);
			}
			if((at & BookHelper.BORDER_EDGE_RIGHT)!=0){
				if(hasBorder)
					matcher.setRightBorderColor(colorHtml);
				else
					matcher.removeRightBorderColor();
				
				matcher.setBorderRight(lineStyle);
			}
			if((at & BookHelper.BORDER_EDGE_BOTTOM)!=0){
				if(hasBorder)
					matcher.setBottomBorderColor(colorHtml);
				else
					matcher.removeBottomBorderColor();
				
				matcher.setBorderBottom(lineStyle);
			}
			style = findStyle(book, matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			if((at & BookHelper.BORDER_EDGE_LEFT)!=0) {
				if(hasBorder)
					BookHelper.setLeftBorderColor(style, color);
				style.setBorderLeft(lineStyle);
			}
			if((at & BookHelper.BORDER_EDGE_TOP)!=0){
				if(hasBorder)
					BookHelper.setTopBorderColor(style, color);
				style.setBorderTop(lineStyle);
			}
			if((at & BookHelper.BORDER_EDGE_RIGHT)!=0){
				if(hasBorder)
					BookHelper.setRightBorderColor(style, color);
				style.setBorderRight(lineStyle);
			}
			if((at & BookHelper.BORDER_EDGE_BOTTOM)!=0){
				if(hasBorder)
					BookHelper.setBottomBorderColor(style, color);
				style.setBorderBottom(lineStyle);
			}
		}
		
		cell.setCellStyle(style);
	}
	
	private static void debugStyle(String msg,int row, int col, Workbook book, NCellStyle style){
		StringBuilder sb = new StringBuilder(msg);
		sb.append("[").append(Ranges.getCellRefString(row, col)).append("]");
		sb.append("Top:[").append(style.getBorderTop()).append(":").append(BookHelper.colorToBorderHTML(book,style.getTopBorderColorColor())).append("]");
		sb.append("Left:[").append(style.getBorderLeft()).append(":").append(BookHelper.colorToBorderHTML(book,style.getLeftBorderColorColor())).append("]");
		sb.append("Bottom:[").append(style.getBorderBottom()).append(":").append(BookHelper.colorToBorderHTML(book,style.getBottomBorderColorColor())).append("]");
		sb.append("Right:[").append(style.getBorderRight()).append(":").append(BookHelper.colorToBorderHTML(book,style.getRightBorderColorColor())).append("]");
		System.out.println(">>"+sb.toString());
	}
	
	public static void setFontBoldWeight(NSheet sheet,int row,int col,short boldWeight,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short fontIdx = orgStyle.getFontIndex();
		final NFont font = book.getFontAt(fontIdx);
		final short orgBoldWeight = font.getBoldweight();
		if (orgBoldWeight == boldWeight) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		final NColor color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Object[] newFont = BookHelper.findOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		
		NCellStyle style = null;
		if(Boolean.FALSE.equals(newFont[1])){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
			matcher.setFontIndex(((Font)newFont[0]).getIndex());
			style = findStyle(sheet.getBook(), matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(((Font)newFont[0]));
		}
		cell.setCellStyle(style);
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setFontItalic(NSheet sheet, int row, int col, boolean italic,HashMap<Integer,CellStyle> cache) {
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short fontIdx = orgStyle.getFontIndex();
		final NFont font = book.getFontAt(fontIdx);
		final boolean orgItalic = font.getItalic();
		if (orgItalic == italic) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		final short boldWeight = font.getBoldweight();
		final NColor color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final byte underline = font.getUnderline();
		final Object[] newFont = BookHelper.findOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		
		NCellStyle style = null;
		if(Boolean.FALSE.equals(newFont[1])){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
			matcher.setFontIndex(((Font)newFont[0]).getIndex());
			style = findStyle(sheet.getBook(), matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(((Font)newFont[0]));
		}
		cell.setCellStyle(style);
		
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setFontUnderline(NSheet sheet,int row,int col, byte underline,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final XBook book = (XBook) sheet.getWorkbook();
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short fontIdx = orgStyle.getFontIndex();
		final NFont font = book.getFontAt(fontIdx);
		final byte orgUnderline = font.getUnderline();
		if (orgUnderline == underline) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		final short boldWeight = font.getBoldweight();
		final NColor color = BookHelper.getFontColor(book, font);
		final short fontHeight = font.getFontHeight();
		final String name = font.getFontName();
		final boolean italic = font.getItalic();
		final boolean strikeout = font.getStrikeout();
		final short typeOffset = font.getTypeOffset();
		final Object[] newFont = BookHelper.findOrCreateFont(book, boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		
		NCellStyle style = null;
		if(Boolean.FALSE.equals(newFont[1])){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
			matcher.setFontIndex(((Font)newFont[0]).getIndex());
			style = findStyle(sheet.getBook(), matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(((Font)newFont[0]));
		}
		cell.setCellStyle(style);
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setTextHAlign(NSheet sheet,int row,int col, short align,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short orgAlign = orgStyle.getAlignment();
		if (align == orgAlign) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
		matcher.setAlignment(align);
		NCellStyle style = findStyle(sheet.getBook(), matcher);
		if(style==null){
			style = cloneCellStyle(cell);
			style.setAlignment(align);
		}
		cell.setCellStyle(style);
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	
	public static void setTextVAlign(NSheet sheet,int row,int col, short valign,HashMap<Integer,CellStyle> cache){
		final NCell cell = XUtils.getOrCreateCell(sheet,row,col);
		final NCellStyle orgStyle = NStyles.getCellStyle(cell);
		final short orgValign = orgStyle.getVerticalAlignment();
		if (valign == orgValign) { //no change, skip
			return;
		}
		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
		if(hitStyle!=null){
			cell.setCellStyle(hitStyle);
			return;
		}
		CellStyleMatcher matcher = new CellStyleMatcher(sheet.getBook(),NStyles.getCellStyle(cell));
		matcher.setVerticalAlignment(valign);
		NCellStyle style = findStyle(sheet.getBook(), matcher);
		if(style==null){
			style = cloneCellStyle(cell);
			style.setVerticalAlignment(valign);
		}
		cell.setCellStyle(style);
		if(cache!=null){
			cache.put((int)orgStyle.getIndex(), style);
		}
	}
	*/
	public static void setDataFormat(NSheet sheet, int row, int col, String format/*,HashMap<Integer,CellStyle> cache*/) {
		final NBook book = sheet.getBook();
		final NCell cell = sheet.getCell(row,col);
		final NCellStyle orgStyle = cell.getCellStyle();
		final String orgFormat = orgStyle.getDataFormat();
		if (format == orgFormat || (format!=null && format.equals(orgFormat))) { //no change, skip
			return;
		}
//		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
//		if(hitStyle!=null){
//			cell.setCellStyle(hitStyle);
//			return;
//		}
//		
		CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);

		matcher.setDataFormat(format);
		NCellStyle style = book.searchCellStyle(matcher);
		if(style==null){
			style = cloneCellStyle(cell);
			style.setDataFormat(format);
		}
		cell.setCellStyle(style);
//		if(cache!=null){
//			cache.put((int)orgStyle.getIndex(), style);
//		}
		
	}
	/*
	
	public static boolean isDefaultStyle(Sheet sheet, NCellStyle style){
		
//		NCellStyle defaultStyle = sheet.getBook().getCellStyleAt((short)0);
//		return defaultStyle.equals(style);
		
		//in xssf, it use toString.equals() which is performance is bad
		//Instated, I just check it's index is 0
		return style!=null && style.getIndex() == 0;
	}
	
	public static NCellStyle getCellStyle(NCell cell){
		NSheet sheet = (NSheet)cell.getSheet();
		//current poi implementation always return default style if the cell doesn't has style
		NCellStyle cellStyle = cell.getCellStyle();
		
		if(cellStyle==null || isDefaultStyle(sheet,cellStyle)){
			cellStyle = cell.getRow().getRowStyle();
		}
		if(cellStyle==null || isDefaultStyle(sheet,cellStyle)){
			cellStyle = sheet.getColumnStyle(cell.getColumnIndex());
		}
		return cellStyle!=null?cellStyle: sheet.getBook().getCellStyleAt((short)0);
	}
	public static NCellStyle getCellStyle(Sheet sheet, int row, int column){
		Row rowObj = sheet.getRow(row);
		NCellStyle cellStyle = null;
		if(rowObj!=null){
			NCell cell = rowObj.getCell(column);
			if(cell!=null){
				cellStyle = cell.getCellStyle(); 
			}
			if(cellStyle==null || isDefaultStyle(sheet,cellStyle)){
				cellStyle = rowObj.getRowStyle();
			}
		}
		if(cellStyle==null || isDefaultStyle(sheet,cellStyle)){
			cellStyle = sheet.getColumnStyle(column);
		}
		
		return cellStyle!=null?cellStyle: sheet.getWorkbook().getCellStyleAt((short)0);
	}
	*/
}
