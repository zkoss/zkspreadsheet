/* WholeStyleUtil.java

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
package org.zkoss.zss.range.impl;

import java.util.Iterator;

import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyleHolder;
import org.zkoss.zss.model.SColumn;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;

/**
 * A utility class to help spreadsheet set style of a row, column and cell
 * @author Dennis.Chen
 * @since 3.5.0
 */
public class WholeStyleUtil {

	public interface StyleApplyer {
		public void applyStyle(SCellStyleHolder holder);
	};
	
	public static void setWholeStyle(SRange range, StyleApplyer applyer){
		if(range.isWholeSheet()){
			//1. it is not possible to set style to all row or columns
			//2. we don't have the way to replace defualt style in component.
			//caller should consider to set cell style to whole column (style on column) and give column from 0 to 1000, because of 
			//in general case we have much less columns then rows. 
			 throw new InvalidModelOpException("don't allow to set style to whole sheet");
		}else if(range.isWholeRow()) {
			setWholeRowCellStyle(range,applyer);
		}else if(range.isWholeColumn()){
			setWholeColumnCellStyle(range,applyer);
		}else{
			for(int r = range.getRow(); r <= range.getLastRow(); r++){
				for (int c = range.getColumn(); c <= range.getLastColumn(); c++){
					SCell cell = range.getSheet().getCell(r,c);
					applyer.applyStyle(cell);
				}
			}
		}
	}
	
	public static void setWholeRowCellStyle(SRange range, StyleApplyer applyer){
		SSheet sheet = range.getSheet();
		for(int r = range.getRow(); r <= range.getLastRow(); r++){
			SRow row = sheet.getRow(r);
			applyer.applyStyle(row);
			
			Iterator<SCell> cells = sheet.getCellIterator(r);
			while(cells.hasNext()){
				SCell cell = cells.next();
				//the case the cell or column has local style
				if(cell.getCellStyle(true)!=null ||
						sheet.getColumn(cell.getColumnIndex()).getCellStyle(true)!=null){
					applyer.applyStyle(cell);
				}
			}
		}
	}
	
	public static void setWholeColumnCellStyle(SRange range, StyleApplyer applyer){
		SSheet sheet = range.getSheet();
		for (int c = range.getColumn(); c <= range.getLastColumn(); c++){
			SColumn column = sheet.getColumn(c);
			applyer.applyStyle(column);
		}
		Iterator<SRow> rows = sheet.getRowIterator();
		while(rows.hasNext()){
			SRow row = rows.next();
			for (int c = range.getColumn(); c <= range.getLastColumn(); c++){
				SCell cell = sheet.getCell(row.getIndex(),c);
				//the case the cell or column has local style
				if(cell.getCellStyle(true)!=null || row.getCellStyle(true)!=null){
					applyer.applyStyle(cell);
				}
			}
		}
	}

	public static void setFillColor(final SRange wholeRange, final String htmlColor) {
		setWholeStyle(wholeRange,new StyleApplyer(){
			@Override
			public void applyStyle(SCellStyleHolder holder) {
				StyleUtil.setFillColor(wholeRange.getSheet().getBook(), holder, htmlColor);
			}});
	}

	public static void setTextHAlign(final SRange wholeRange,
			final org.zkoss.zss.model.SCellStyle.Alignment hAlignment) {
		setWholeStyle(wholeRange,new StyleApplyer(){
			@Override
			public void applyStyle(SCellStyleHolder holder) {
				StyleUtil.setTextHAlign(wholeRange.getSheet().getBook(), holder, hAlignment);
			}});
	}

	public static void setTextVAlign(final SRange wholeRange, 
			final org.zkoss.zss.model.SCellStyle.VerticalAlignment vAlignment) {
		setWholeStyle(wholeRange,new StyleApplyer(){
			@Override
			public void applyStyle(SCellStyleHolder holder) {
				StyleUtil.setTextVAlign(wholeRange.getSheet().getBook(), holder, vAlignment);
			}});
	}

	public static void setTextWrap(final SRange wholeRange, final boolean wraptext) {
		setWholeStyle(wholeRange,new StyleApplyer(){
			@Override
			public void applyStyle(SCellStyleHolder holder) {
				StyleUtil.setTextWrap(wholeRange.getSheet().getBook(), holder, wraptext);
			}});
	}

	public static void setDataFormat(final SRange wholeRange, final String format) {
		setWholeStyle(wholeRange,new StyleApplyer(){
			@Override
			public void applyStyle(SCellStyleHolder holder) {
				StyleUtil.setDataFormat(wholeRange.getSheet().getBook(), holder, format);
			}});
	}	
}
