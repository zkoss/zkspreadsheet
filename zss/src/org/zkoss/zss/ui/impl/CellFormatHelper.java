/* CellFormatHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 29, 2008 11:14:44 AM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
 */
package org.zkoss.zss.ui.impl;

import java.awt.Color;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Rect;

/**
 * @author Dennis.Chen
 * 
 */
public class CellFormatHelper {

	/*
	 * cell to get the format, could be null.
	 */
	private Cell _cell;

	private Sheet _sheet;
	
	private Book _book;

	private int _row;

	private int _col;

	private boolean hasRightBorder_set = false;
	private boolean hasRightBorder = false;
	
	private MergeMatrixHelper _mmHelper;

	public CellFormatHelper(Sheet sheet, int row, int col, MergeMatrixHelper mmhelper) {
		_sheet = sheet;
		_book = (Book) _sheet.getWorkbook(); 
		_row = row;
		_col = col;
		_cell = Utils.getCell(sheet, row, col);
		_mmHelper = mmhelper;
	}

	public String getHtmlStyle() {

		StringBuffer sb = new StringBuffer();
		if (_cell != null) {
			final FormatText ft = Utils.getFormatText(_cell);
			final boolean isRichText = ft.isRichTextString();
			final RichTextString rstr = isRichText ? ft.getRichTextString() : null;
			final String txt = rstr != null ? rstr.getString() : ft.getCellFormatResult().text;
			final CellStyle style = _cell.getCellStyle();
			
			if (style == null)
				return "";

			String bgColor = BookHelper.indexToRGB(_book, style.getFillForegroundColor());
			if (BookHelper.AUTO_COLOR.equals(bgColor)) {
				bgColor = null;
			}
			if (bgColor != null) {
				sb.append("background-color:").append(bgColor).append(";");
			}
			if (!isRichText && ft.getCellFormatResult().textColor != null) {
				final Color textColor = ft.getCellFormatResult().textColor;
				final String htmlColor = toHTMLColor(textColor);
				sb.append("color:").append(htmlColor).append(";");
			}

			if (bgColor != null && (txt == null || txt.trim().equals(""))) {
				// not text but has bg color, i must set the z-index to 0
				// otherwise, it will cover the overflow text of previous cell
				// use 0, safari not wrok when zindex = -1;
				sb.append("z-index:0;");
			}
		}

		processBottomBorder(sb);
		processRightBorder(sb);

		return sb.toString();
	}

	private String toHTMLColor(Color color) {
		final int r = color.getRed();
		final int g = color.getGreen();
		final int b = color.getBlue();
		return "#"+BookHelper.toHex(r)+BookHelper.toHex(g)+BookHelper.toHex(b);
	}
	
	private boolean processBottomBorder(StringBuffer sb) {

		boolean hitBottom = false;

		if (_cell != null) {
			CellStyle style = _cell.getCellStyle();
			
			
			if (style != null){
				int bb = style.getBorderBottom();
				String color = BookHelper.indexToRGB(_book, style.getBottomBorderColor());
				hitBottom = appendBorderStyle(sb, "bottom", bb, color);
			}
		}
		Cell next = null;
		if (!hitBottom) {
			next = Utils.getCell(_sheet, _row + 1, _col);
			/*if(next == null){ // don't search into merge ranges
				//check is _row+1,_col in merge range
				MergedRect rect = _mmHelper.getMergeRange(_row+1, _col);
				if(rect !=null){
					next = _sheet.getCell(rect.getTop(),rect.getLeft());
				}
			}*/
			if (next != null) {
				CellStyle style = next.getCellStyle();
				if (style != null){
					int bb = style.getBorderTop();// get top border of
					String color = BookHelper.indexToRGB(_book, style.getTopBorderColor());
					// set next row top border as cell's bottom border;
					hitBottom = appendBorderStyle(sb, "bottom", bb, color);
				}
			}
		}
		
		//border depends on next cell's background color
		if(!hitBottom && next !=null){
			CellStyle style = next.getCellStyle();
			if (style != null){
				String bgColor = BookHelper.indexToRGB(_book, style.getFillBackgroundColor());
				if (BookHelper.AUTO_COLOR.equals(bgColor)) {
					bgColor = null;
				}
				if (bgColor != null) {
					hitBottom = appendBorderStyle(sb, "bottom", CellStyle.BORDER_THIN, bgColor);
				}
			}
		}
		
		//border depends on current cell's background color
		if(!hitBottom && _cell !=null){
			CellStyle style = _cell.getCellStyle();
			
			if (style != null){
				String bgColor = BookHelper.indexToRGB(_book, style.getFillBackgroundColor());
				if (BookHelper.AUTO_COLOR.equals(bgColor)) {
					bgColor = null;
				}
				if (bgColor != null) {
					hitBottom = appendBorderStyle(sb, "bottom", CellStyle.BORDER_THIN, bgColor);
				}
			}
		}
		
		return hitBottom;
	}

	private boolean processRightBorder(StringBuffer sb) {
		boolean hitRight = false;
		MergedRect rect=null;
		boolean hitMerge = false;
		//find right border of target cell 
		if (_cell != null) {
			Cell right = _cell;
			rect = _mmHelper.getMergeRange(_row, _col);
			if(rect!=null){
				hitMerge = true;
				right = Utils.getCell(_sheet, _row, rect.getRight());
			}
			CellStyle style = right.getCellStyle();
			if (style != null){
				int bb = style.getBorderRight();
				String color = BookHelper.indexToRGB(_book, style.getRightBorderColor());
				hitRight = appendBorderStyle(sb, "right", bb, color);
			}
		}

		Cell next = null;
		//if no border for target cell,then check is this cell in a merge range
		//if(true) then try to get next cell after this merge range
		//else get next cell of this cell
		if(!hitRight){
			int c = hitMerge?rect.getRight()+1:_col+1;
			next = Utils.getCell(_sheet, _row, c);
			//find the right cell of merge range.
			if(next!=null){
				CellStyle style = next.getCellStyle();
				if (style != null){
					int bb = style.getBorderLeft();//get left here
					String color = BookHelper.indexToRGB(_book, style.getLeftBorderColor());
						hitRight = appendBorderStyle(sb, "right", bb, color);
				}
			}
		}

		//border depends on next cell's background color
		if(!hitRight && next !=null){
			CellStyle style = next.getCellStyle();
			if (style != null){
				String bgColor = BookHelper.indexToRGB(_book, style.getFillBackgroundColor());
				if (BookHelper.AUTO_COLOR.equals(bgColor)) {
					bgColor = null;
				}
				
				if (bgColor != null) {
					hitRight = appendBorderStyle(sb, "right", CellStyle.BORDER_THIN, bgColor);
				}
			}
		}
		//border depends on current cell's background color
		if(!hitRight && _cell !=null){
			CellStyle style = _cell.getCellStyle();
			if (style != null){
				String bgColor = BookHelper.indexToRGB(_book, style.getFillBackgroundColor());
				if (BookHelper.AUTO_COLOR.equals(bgColor)) {
					bgColor = null;
				}
				if (bgColor != null) {
					hitRight = appendBorderStyle(sb, "right", CellStyle.BORDER_THIN, bgColor);
				}
			}
		}
		
		return hitRight;
	}

	private boolean appendBorderStyle(StringBuffer sb, String locate, int bs, String color) {
		if (bs == CellStyle.BORDER_NONE)
			return false;
		
		sb.append("border-").append(locate).append(":");
		switch(bs) {
		case CellStyle.BORDER_DASHED:
		case CellStyle.BORDER_DOTTED:
			sb.append("dashed");
			break;
		case CellStyle.BORDER_HAIR:
			sb.append("dotted");
			break;
		default:
			sb.append("solid");
		}
		sb.append(" 1px");

		if (color != null) {
			if (BookHelper.AUTO_COLOR.equals(color)) {
				color = "#000000";
			}
			sb.append(" ");
			sb.append(color);
		}

		sb.append(";");
		return true;
	}

	public String getInnerHtmlStyle() {
		if (_cell != null) {
			CellStyle style = _cell.getCellStyle();
			if (style == null)
				return "";

			final StringBuffer sb = new StringBuffer();
			sb.append(BookHelper.getTextCSSStyle(_book, _cell));
			
			final Font font = _book.getFontAt(style.getFontIndex());
			
			sb.append(BookHelper.getFontCSSStyle(_book, font));
			return sb.toString();
		}
		return "";
	}

	public boolean hasRightBorder() {
		if(hasRightBorder_set){
			return hasRightBorder;
		}else{
			hasRightBorder = processRightBorder(new StringBuffer());
			hasRightBorder_set = true;
		}
		return hasRightBorder;
	}

}
