/* UtilFns.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 17, 2007 5:12:20 PM     2007, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.fn;

import java.io.IOException;
import java.io.StringWriter;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.zk.ui.UiException;
//import org.zkoss.zss.model.Cell;
//import org.zkoss.zss.model.Format;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;


/**
 * 
 * This class is for Spreadsheet Taglib use only, don't use it as a utility .
 * @author Dennis.Chen
 *
 */
public class UtilFns {

	/**
	 * Gets Column name of a sheet
	 */
	static public String getColumntitle(Spreadsheet ss,int index){
		return ss.getColumntitle(index);
	}
	
	
	/**
	 * Gets Row name of a sheet
	 */
	static public String getRowtitle(Spreadsheet ss,int index){
		return ss.getRowtitle(index);
	}
	
	/**
	 * Gets Cell text by given row and column
	 */
	static public String getCelltext(Spreadsheet ss,int row,int column){
		/*List list = ss.getBook().getSheets();
		if(list.size()<=ss.getSelectedIndex()){
			throw new XelException("No such sheet :"+ss.getSelectedIndex());
		}*/
		Worksheet sheet = ss.getSelectedSheet();
		final Cell cell = Utils.getCell(sheet, row, column);
		String text = "";
		if (cell != null) {
			boolean wrap = cell.getCellStyle().getWrapText();
			final FormatText ft = Utils.getFormatText(cell);
			if (ft != null) {
				if (ft.isRichTextString()) {
					final RichTextString rstr = ft.getRichTextString();
					text = rstr == null ? "" : Utils.formatRichTextString(sheet, rstr, wrap);
				} else if (ft.isCellFormatResult()) {
					text = Utils.escapeCellText(ft.getCellFormatResult().text, wrap, true);
				}
			}
			final Hyperlink hlink = Utils.getHyperlink(cell);
			if (hlink != null) {
				text = Utils.formatHyperlink(sheet, hlink, text, wrap);
			}
		}
		return text;
	}
	
	//Gets Cell edit text by given row and column
	static public String getEdittext(Spreadsheet ss,int row,int column){
		Worksheet sheet = ss.getSelectedSheet();
		final Cell cell = Utils.getCell(sheet, row, column);
		return cell != null ? Utils.getEditText(cell) : "";
	}
	
	static public Integer getRowBegin(Spreadsheet ss){
		return Integer.valueOf(0);
	}
	static public Integer getRowEnd(Spreadsheet ss){
		int max = ss.getMaxrows();
		max = max<=20?max-1:20;
		return Integer.valueOf(max);//Integer.valueOf(ss.getMaxrow()-1);
	}
	static public Integer getColBegin(Spreadsheet ss){
		return Integer.valueOf(0);
	}
	static public Integer getColEnd(Spreadsheet ss){
		int max = ss.getMaxcolumns();
		
		max = max<=10?max-1:10;
		
		int row_top = getRowBegin(ss).intValue();
		int row_bottom = getRowEnd(ss).intValue();
		
		Worksheet sheet = ss.getSelectedSheet();
		MergeMatrixHelper mmhelper = ((SpreadsheetCtrl)ss.getExtraCtrl()).getMergeMatrixHelper(sheet);
		Set blocks = mmhelper.getRangesByColumn(max);
		Iterator iter = blocks.iterator();
		while(iter.hasNext()){
			Rect rect = (Rect)iter.next();
			int top = rect.getTop();
			//int left = rect.getLeft();
			int right = rect.getRight();
			//int bottom = rect.getBottom();
			if(top<row_top || top<row_bottom){
				continue;
			}
			
			if(max<right){
				max = right;
			}
		}
		
		return Integer.valueOf(max);//Integer.valueOf(ss.getMaxcolumn()-1);
	}
	
	static public String getRowOuterAttrs(Spreadsheet ss,int row){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getRowOuterAttrs(row);
	}
	
	static public String getCellOuterAttrs(Spreadsheet ss,int row,int col){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getCellOuterAttrs(row,col);
	}
	
	static public String getCellInnerAttrs(Spreadsheet ss,int row,int col){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getCellInnerAttrs(row,col);
	}
	static public String getTopHeaderOuterAttrs(Spreadsheet ss,int col){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getTopHeaderOuterAttrs(col);
	}
	static public String getTopHeaderInnerAttrs(Spreadsheet ss,int col){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getTopHeaderInnerAttrs(col);
	}
	
	static public String getLeftHeaderOuterAttrs(Spreadsheet ss,int row){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getLeftHeaderOuterAttrs(row);
	}
	
	static public String getLeftHeaderInnerAttrs(Spreadsheet ss,int row){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getLeftHeaderInnerAttrs(row);
	}
	
	static public Boolean getTopHeaderHiddens(Spreadsheet ss, int col) {
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getTopHeaderHiddens(col);
	}
	
	static public Boolean getLeftHeaderHiddens(Spreadsheet ss,int row){
		return ((SpreadsheetCtrl)ss.getExtraCtrl()).getLeftHeaderHiddens(row);
	}
}
