/* ReserveUtil.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/7/25, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
 */
package org.zkoss.zss.undo.impl;

import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;

/**
 * 
 * @author dennis
 * 
 */
public class ReserveUtil {
//	public enum ReserveType {
//		ALL, STYLE, DATA
//	}
	public static final int RESERVE_DATA = 1;
	public static final int RESERVE_STYLE = 2;
	public static final int RESERVE_MERGE = 4;
	public static final int RESERVE_ALL = RESERVE_DATA|RESERVE_STYLE|RESERVE_MERGE;
	public static ReservedResult reserve(Sheet sheet, int row, int column,
			int lastRow, int lastColumn, int reserveType) {

		ReservedResult result = new ReservedResult(sheet,row,column,lastRow,lastColumn);

		ReservedCellContent[][] data = null;
		CellStyle[][] styles = null;
		Rect[] mergeInfo = null;
		if((reserveType & RESERVE_DATA)!=0){
			result.setContent(data = new ReservedCellContent[lastRow - row + 1][lastColumn
			                                          					- column + 1]);
		}
		if((reserveType & RESERVE_STYLE)!=0){
			result.setStyles(styles = new CellStyle[lastRow - row + 1][lastColumn
			                                       					- column + 1]);
		}
		if((reserveType & RESERVE_MERGE)!=0){
			
			
		}
		

		for (int i = row; i <= lastRow; i++) {
			for (int j = column; j <= lastColumn; j++) {
				Range r = Ranges.range(sheet, i, j);
				if (styles != null) {
					styles[i - row][j - column] = r.getCellStyle();
				}
				if (data != null) {
					data[i - row][j - column] = ReservedCellContent.reserve(r);
				}
			}
		}

		return result;
	}

	public static class ReservedResult {
		final Sheet _sheet;
		ReservedCellContent[][] _content = null;
		CellStyle[][] _styles = null;
		Rect[] _mergeInfo;
		final int _row, _column, _lastRow, _lastColumn;

		public ReservedResult(Sheet sheet, int row, int column, int lastRow,
				int lastColumn) {
			_sheet = sheet;
			_row = row;
			_column = column;
			_lastRow = lastRow;
			_lastColumn = lastColumn;
		}

		public ReservedCellContent[][] getContent() {
			return _content;
		}

		private void setContent(ReservedCellContent[][] data) {
			this._content = data;
		}

		public CellStyle[][] getStyles() {
			return _styles;
		}

		private void setStyles(CellStyle[][] styles) {
			this._styles = styles;
		}

		public Sheet getSheet() {
			return _sheet;
		}

		public int getRow() {
			return _row;
		}

		public int getColumn() {
			return _column;
		}

		public int getLastRow() {
			return _lastRow;
		}

		public int getLastColumn() {
			return _lastColumn;
		}
		
		public Rect[] getMergeInfo() {
			return _mergeInfo;
		}

		public void setMergeInfo(Rect[] mergeInfo) {
			this._mergeInfo = mergeInfo;
		}

		public void restore(){
			int row = getRow();
			int column = getColumn();
			int lastRow = getLastRow();
			int lastColumn = getLastColumn();
			Sheet sheet = getSheet();
			
			ReservedCellContent[][] data = getContent();
			CellStyle[][] styles = getStyles();
			Rect[] mergeInfo = getMergeInfo(); 
			
			
			if(mergeInfo!=null){
				for(Rect rect:mergeInfo){
					//re merge back
					Range r = Ranges.range(sheet,rect);
					r.merge(false);
				}
			}
			
			for(int i=row;i<=lastRow;i++){
				for(int j=column;j<=lastColumn;j++){
					Range r = Ranges.range(sheet,i,j);
					if(styles!=null){
						r.setCellStyle(styles[i-row][j-column]);
					}
					if(data!=null){
						data[i-row][j-column].apply(r);
					}
				}
			}

		}
	}
	
	public static class ReservedCellContent {

		boolean _blank = false;;
		String _editText;
		
		public ReservedCellContent(){
			this._blank = true;
		}
		
		public ReservedCellContent(String editText){
			this._editText = editText;
		}
		
		public void apply(Range range){
			if(_blank){
				range.clearContents();
			}else{
				try{
					range.setCellEditText(_editText);
				}catch(IllegalFormulaException x){};//eat in this mode
			}
		}
		
		public static ReservedCellContent reserve(Range range){
			CellData d = range.getCellData();
			
			if(d.isBlank()){
				return new ReservedCellContent();
			}
			
			String editText = d.getEditText();
			//TODO handle other data someday(hyperlink, comment)
			return new ReservedCellContent(editText);
		}
	}

}
