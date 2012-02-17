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

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;

/**
 * @author Dennis.Chen
 * 
 */
public class CellFormatHelper {

	/*
	 * cell to get the format, could be null.
	 */
	private Cell _cell;

	private Worksheet _sheet;
	
	private Book _book;

	private int _row;

	private int _col;

	private boolean hasRightBorder_set = false;
	private boolean hasRightBorder = false;
	
	private MergeMatrixHelper _mmHelper;

	public CellFormatHelper(Worksheet sheet, int row, int col, MergeMatrixHelper mmhelper) {
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
			final CellStyle style = _cell.getCellStyle();
			
			if (style == null)
				return "";

				
			
			//String bgColor = BookHelper.indexToRGB(_book, style.getFillForegroundColor());
			//ZSS-34 cell background color does not show in excel
			//20110819, henrichen: if fill pattern is NO_FILL, shall not show the cell background color
			String bgColor = style.getFillPattern() != CellStyle.NO_FILL ? 
				BookHelper.colorToHTML(_book, style.getFillForegroundColorColor()) : null;
			if (BookHelper.AUTO_COLOR.equals(bgColor)) {
				bgColor = null;
			}
			if (bgColor != null) {
				sb.append("background-color:").append(bgColor).append(";");
			}
			final FormatText ft = Utils.getFormatText(_cell);
			final boolean isRichText = ft.isRichTextString();
			final RichTextString rstr = isRichText ? ft.getRichTextString() : null;
			final String txt = rstr != null ? rstr.getString() : ft.getCellFormatResult().text;
			
			if(_cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				sb.append("z-index:-1;"); //For IE6/IE7's overflow
			}
		}

		processBottomBorder(sb);
		processRightBorder(sb);

		return sb.toString();
	}

	private String toHTMLColor(Color color) {
		return BookHelper.awtColorToHTMLColor(color);
	}
	
	private boolean processBottomBorder(StringBuffer sb) {

		boolean hitBottom = false;

		if (_cell != null) {
			CellStyle style = _cell.getCellStyle();
			
			
			if (style != null){
				int bb = style.getBorderBottom();
				//String color = BookHelper.indexToRGB(_book, style.getBottomBorderColor());
				String color = BookHelper.colorToHTML(_book, style.getBottomBorderColorColor());
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
					//String color = BookHelper.indexToRGB(_book, style.getTopBorderColor());
					String color = BookHelper.colorToHTML(_book, style.getTopBorderColorColor());
					// set next row top border as cell's bottom border;
					hitBottom = appendBorderStyle(sb, "bottom", bb, color);
				}
			}
		}
		
		//border depends on next cell's background color
		if(!hitBottom && next !=null){
			CellStyle style = next.getCellStyle();
			if (style != null){
				//String bgColor = BookHelper.indexToRGB(_book, style.getFillForegroundColor());
				//ZSS-34 cell background color does not show in excel
				String bgColor = style.getFillPattern() != CellStyle.NO_FILL ? 
						BookHelper.colorToHTML(_book, style.getFillForegroundColorColor()) : null;
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
				//String bgColor = BookHelper.indexToRGB(_book, style.getFillForegroundColor());
				//ZSS-34 cell background color does not show in excel
				String bgColor = style.getFillPattern() != CellStyle.NO_FILL ? 
						BookHelper.colorToHTML(_book, style.getFillForegroundColorColor()) : null;
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
			if (right != null) {
				CellStyle style = right.getCellStyle();
				if (style != null){
					int bb = style.getBorderRight();
					//String color = BookHelper.indexToRGB(_book, style.getRightBorderColor());
					String color = BookHelper.colorToHTML(_book, style.getRightBorderColorColor());
					hitRight = appendBorderStyle(sb, "right", bb, color);
				}
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
					//String color = BookHelper.indexToRGB(_book, style.getLeftBorderColor());
					//ZSS-34 cell background color does not show in excel
					String color = style.getFillPattern() != CellStyle.NO_FILL ? 
							BookHelper.colorToHTML(_book, style.getLeftBorderColorColor()) : null;
						hitRight = appendBorderStyle(sb, "right", bb, color);
				}
			}
		}

		//border depends on next cell's background color
		if(!hitRight && next !=null){
			CellStyle style = next.getCellStyle();
			if (style != null){
				//String bgColor = BookHelper.indexToRGB(_book, style.getFillForegroundColor());
				//ZSS-34 cell background color does not show in excel
				String bgColor = style.getFillPattern() != CellStyle.NO_FILL ? 
						BookHelper.colorToHTML(_book, style.getFillForegroundColorColor()) : null;
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
				//String bgColor = BookHelper.indexToRGB(_book, style.getFillForegroundColor());
				//ZSS-34 cell background color does not show in excel
				String bgColor = style.getFillPattern() != CellStyle.NO_FILL ? 
						BookHelper.colorToHTML(_book, style.getFillForegroundColorColor()) : null;
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
			
			//vertical alignment
			int verticalAlignment = style.getVerticalAlignment();
			sb.append("display: table-cell;");
			switch (verticalAlignment) {
			case CellStyle.VERTICAL_TOP:
				sb.append("vertical-align: top;");
				break;
			case CellStyle.VERTICAL_CENTER:
				sb.append("vertical-align: middle;");
				break;
			case CellStyle.VERTICAL_BOTTOM:
				sb.append("vertical-align: bottom;");
				break;
			}
			
			final Font font = _book.getFontAt(style.getFontIndex());
			
			//sb.append(BookHelper.getFontCSSStyle(_book, font));
			sb.append(BookHelper.getFontCSSStyle(_cell, font));

			//condition color
			final FormatText ft = Utils.getFormatText(_cell);
			final boolean isRichText = ft.isRichTextString();
			if (!isRichText && ft.getCellFormatResult().textColor != null) {
				final Color textColor = ft.getCellFormatResult().textColor;
				final String htmlColor = toHTMLColor(textColor);
				sb.append("color:").append(htmlColor).append(";");
			}

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
