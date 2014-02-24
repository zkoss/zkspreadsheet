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
public class StyleUtil {
//	private static final Log log = Log.lookup(NStyles.class);
	
	public static NCellStyle cloneCellStyle(NCell cell) {
		final NCellStyle destination = cell.getSheet().getBook().createCellStyle(cell.getCellStyle(), true);
		return destination;
	}
	
	public static void setFontColor(NSheet sheet, int row, int col, String color/*,HashMap<Integer,NCellStyle> cache*/){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		final NColor orgColor = orgFont.getColor();
		final NColor newColor = book.createColor(color);
		if (orgColor == newColor || orgColor != null && orgColor.equals(newColor)) {
			return;
		}
		
//		NCellStyle hitStyle = cache==null?null:cache.get((int)orgStyle.getIndex());
//		if(hitStyle!=null){
//			cell.setCellStyle(hitStyle);
//			return;
//		}
		
		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setColor(color);
		
		NFont font = book.searchFont(fontmatcher);
		
		
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setColor(newColor);
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
	
	
	public static void setFillColor(NSheet sheet, int row, int col, String htmlColor){
		final NBook book = sheet.getBook();
		final NCell cell = sheet.getCell(row,col);
		final NCellStyle orgStyle = cell.getCellStyle();
		final NColor orgColor = orgStyle.getFillColor();
		final NColor newColor = book.createColor(htmlColor);
		if (orgColor == newColor || orgColor != null  && orgColor.equals(newColor)) { //no change, skip
			return;
		}
		
		CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
		matcher.setFillColor(htmlColor);
		matcher.setFillPattern(NCellStyle.FillPattern.SOLID_FOREGROUND);
		
		NCellStyle style = book.searchCellStyle(matcher);
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFillColor(newColor);
			style.setFillPattern(NCellStyle.FillPattern.SOLID_FOREGROUND);
		}
		cell.setCellStyle(style);
		
	}
	
	public static void setTextWrap(NSheet sheet,int row,int col,boolean wrap){
		final NBook book = sheet.getBook();
		final NCell cell = sheet.getCell(row,col);
		final NCellStyle orgStyle = cell.getCellStyle();
		final boolean textWrap = orgStyle.isWrapText();
		if (wrap == textWrap) { //no change, skip
			return;
		}
		
		CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
		matcher.setWrapText(wrap);
		NCellStyle style = book.searchCellStyle(matcher);
		if(style==null){
			style  = cloneCellStyle(cell);
			style.setWrapText(wrap);
		}
		cell.setCellStyle(style);
	}
	
	public static void setFontHeightPoints(NSheet sheet,int row,int col,int fontHeightPoints){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		
		final int orgSize = orgFont.getHeightPoints();
		if (orgSize == fontHeightPoints) { //no change, skip
			return;
		}
		
		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setHeightPoints(fontHeightPoints);
		
		NFont font = book.searchFont(fontmatcher);
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setHeightPoints(fontHeightPoints);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
	}
	
	public static void setFontStrikethrough(NSheet sheet,int row,int col, boolean strikeout){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		
		final boolean orgStrikeout = orgFont.isStrikeout();
		if (orgStrikeout == strikeout) { //no change, skip
			return;
		}

		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setStrikeout(strikeout);
		
		NFont font = book.searchFont(fontmatcher);
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setStrikeout(strikeout);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
		
	}
	
	public static void setFontName(NSheet sheet,int row,int col,String name){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		
		final String orgName = orgFont.getName();
		if (orgName.equals(name)) { //no change, skip
			return;
		}
		
		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setName(name);
		
		NFont font = book.searchFont(fontmatcher);
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setName(name);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
		
	}
	
	public static final short BORDER_EDGE_BOTTOM		= 0x01;
	public static final short BORDER_EDGE_RIGHT			= 0x02;
	public static final short BORDER_EDGE_TOP			= 0x04;
	public static final short BORDER_EDGE_LEFT			= 0x08;
	public static final short BORDER_EDGE_ALL			= BORDER_EDGE_BOTTOM|BORDER_EDGE_RIGHT|BORDER_EDGE_TOP|BORDER_EDGE_LEFT;
	
	public static void setBorder(NSheet sheet,int row,int col, String color, NCellStyle.BorderType linestyle){
		setBorder(sheet,row,col, color, linestyle, BORDER_EDGE_ALL);
	}
	
	public static void setBorderTop(NSheet sheet,int row,int col,String color, NCellStyle.BorderType linestyle){
		setBorder(sheet,row,col, color, linestyle, BORDER_EDGE_TOP);
	}
	public static void setBorderLeft(NSheet sheet,int row,int col,String color, NCellStyle.BorderType linestyle){
		setBorder(sheet,row,col, color, linestyle, BORDER_EDGE_LEFT);
	}
	public static void setBorderBottom(NSheet sheet,int row,int col,String color, NCellStyle.BorderType linestyle){
		setBorder(sheet,row,col, color, linestyle, BORDER_EDGE_BOTTOM);
	}
	public static void setBorderRight(NSheet sheet,int row,int col,String color, NCellStyle.BorderType linestyle){
		setBorder(sheet,row,col, color, linestyle, BORDER_EDGE_RIGHT);
	}
	
	public static void setBorder(NSheet sheet,int row,int col, String htmlColor, NCellStyle.BorderType lineStyle, short at){
		final NBook book = sheet.getBook();
		final NCell cell = sheet.getCell(row,col);
		final NCellStyle orgStyle = cell.getCellStyle();
		//ZSS-464 try to search existed matched style
		NCellStyle style = null;
		final NColor color = book.createColor(htmlColor);
		boolean hasBorder = lineStyle != NCellStyle.BorderType.NONE;
		if(htmlColor!=null){
			final NCellStyle oldstyle = cell.getCellStyle();
			CellStyleMatcher matcher = new CellStyleMatcher(oldstyle);
			if((at & BORDER_EDGE_LEFT)!=0) {
				if(hasBorder)
					matcher.setBorderLeftColor(htmlColor);
				else
					matcher.removeBorderLeftColor();
				
				matcher.setBorderLeft(lineStyle);
			}
			if((at & BORDER_EDGE_TOP)!=0){
				if(hasBorder) 
					matcher.setBorderTopColor(htmlColor);
				else
					matcher.removeBorderTopColor();
				
				matcher.setBorderTop(lineStyle);
			}
			if((at & BORDER_EDGE_RIGHT)!=0){
				if(hasBorder)
					matcher.setBorderRightColor(htmlColor);
				else
					matcher.removeBorderRightColor();
				
				matcher.setBorderRight(lineStyle);
			}
			if((at & BORDER_EDGE_BOTTOM)!=0){
				if(hasBorder)
					matcher.setBorderBottomColor(htmlColor);
				else
					matcher.removeBorderBottomColor();
				
				matcher.setBorderBottom(lineStyle);
			}
			style = book.searchCellStyle(matcher);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			if((at & BORDER_EDGE_LEFT)!=0) {
				if(hasBorder)
					style.setBorderLeftColor(color);
				style.setBorderLeft(lineStyle);
			}
			if((at & BORDER_EDGE_TOP)!=0){
				if(hasBorder)
					style.setBorderTopColor(color);
				style.setBorderTop(lineStyle);
			}
			if((at & BORDER_EDGE_RIGHT)!=0){
				if(hasBorder)
					style.setBorderRightColor(color);
				style.setBorderRight(lineStyle);
			}
			if((at & BORDER_EDGE_BOTTOM)!=0){
				if(hasBorder)
					style.setBorderBottomColor(color);
				style.setBorderBottom(lineStyle);
			}
		}
		
		cell.setCellStyle(style);
	}
	
//	private static void debugStyle(String msg,int row, int col, Workbook book, NCellStyle style){
//		StringBuilder sb = new StringBuilder(msg);
//		sb.append("[").append(Ranges.getCellRefString(row, col)).append("]");
//		sb.append("Top:[").append(style.getBorderTop()).append(":").append(BookHelper.colorToBorderHTML(book,style.getTopBorderColorColor())).append("]");
//		sb.append("Left:[").append(style.getBorderLeft()).append(":").append(BookHelper.colorToBorderHTML(book,style.getLeftBorderColorColor())).append("]");
//		sb.append("Bottom:[").append(style.getBorderBottom()).append(":").append(BookHelper.colorToBorderHTML(book,style.getBottomBorderColorColor())).append("]");
//		sb.append("Right:[").append(style.getBorderRight()).append(":").append(BookHelper.colorToBorderHTML(book,style.getRightBorderColorColor())).append("]");
//		System.out.println(">>"+sb.toString());
//	}
	
	public static void setFontBoldWeight(NSheet sheet,int row,int col,NFont.Boldweight boldWeight){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		
		final NFont.Boldweight orgBoldWeight = orgFont.getBoldweight();
		if (orgBoldWeight.equals(boldWeight)) { //no change, skip
			return;
		}
		
		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setBoldweight(boldWeight);
		
		NFont font = book.searchFont(fontmatcher);
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setBoldweight(boldWeight);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
	}
	
	public static void setFontItalic(NSheet sheet, int row, int col, boolean italic) {
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		
		final boolean orgItalic = orgFont.isItalic();
		if (orgItalic == italic) { //no change, skip
			return;
		}

		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setItalic(italic);
		
		NFont font = book.searchFont(fontmatcher);
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setItalic(italic);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
		
	}
	
	public static void setFontUnderline(NSheet sheet,int row,int col, NFont.Underline underline){
		final NCell cell = sheet.getCell(row,col);
		final NBook book = sheet.getBook();
		final NCellStyle orgStyle = cell.getCellStyle();
		NFont orgFont = orgStyle.getFont();
		
		final NFont.Underline orgUnderline = orgFont.getUnderline();
		if (orgUnderline.equals(underline)) { //no change, skip
			return;
		}
		
		FontMatcher fontmatcher = new FontMatcher(orgFont);
		fontmatcher.setUnderline(underline);
		
		NFont font = book.searchFont(fontmatcher);
		
		NCellStyle style = null;
		if(font!=null){//search it since we have existed font
			CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
			matcher.setFont(font);
			style = book.searchCellStyle(matcher);
		}else{
			font = book.createFont(orgFont,true);
			font.setUnderline(underline);
		}
		
		if(style==null){
			style = cloneCellStyle(cell);
			style.setFont(font);
		}
		cell.setCellStyle(style);
	}
	
	public static void setTextHAlign(NSheet sheet,int row,int col, NCellStyle.Alignment align){
		final NBook book = sheet.getBook();
		final NCell cell = sheet.getCell(row,col);
		final NCellStyle orgStyle = cell.getCellStyle();
		final NCellStyle.Alignment orgAlign = orgStyle.getAlignment();
		if (align.equals(orgAlign)) { //no change, skip
			return;
		}
		
		CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
		matcher.setAlignment(align);
		NCellStyle style = book.searchCellStyle(matcher);
		if(style==null){
			style = cloneCellStyle(cell);
			style.setAlignment(align);
		}
		cell.setCellStyle(style);
	}
	
	public static void setTextVAlign(NSheet sheet,int row,int col, NCellStyle.VerticalAlignment valign){
		final NBook book = sheet.getBook();
		final NCell cell = sheet.getCell(row,col);
		final NCellStyle orgStyle = cell.getCellStyle();
		final NCellStyle.VerticalAlignment orgValign = orgStyle.getVerticalAlignment();
		if (valign.equals(orgValign)) { //no change, skip
			return;
		}

		CellStyleMatcher matcher = new CellStyleMatcher(orgStyle);
		matcher.setVerticalAlignment(valign);
		NCellStyle style = book.searchCellStyle(matcher);
		if(style==null){
			style = cloneCellStyle(cell);
			style.setVerticalAlignment(valign);
		}
		cell.setCellStyle(style);

	}
	
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
}
