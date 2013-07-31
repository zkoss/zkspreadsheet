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

/**
 * 
 * @author dennis
 * 
 */
public class ReserveUtil {
	public enum ReserveType {
		ALL, STYLE, DATA
	}
	public static ReservedResult reserve(Sheet sheet, int row, int column,
			int lastRow, int lastColumn, ReserveType type) {

		ReservedResult result = new ReservedResult(sheet,row,column,lastRow,lastColumn);

		ReservedCellData[][] data = null;
		CellStyle[][] styles = null;

		switch (type) {
		case DATA:
			result.setData(data = new ReservedCellData[lastRow - row + 1][lastColumn
					- column + 1]);
			break;
		case STYLE:
			result.setStyles(styles = new CellStyle[lastRow - row + 1][lastColumn
					- column + 1]);
			break;
		case ALL:
			result.setData(data = new ReservedCellData[lastRow - row + 1][lastColumn
					- column + 1]);
			result.setStyles(styles = new CellStyle[lastRow - row + 1][lastColumn
					- column + 1]);
		}

		for (int i = row; i <= lastRow; i++) {
			for (int j = column; j <= lastColumn; j++) {
				Range r = Ranges.range(sheet, i, j);
				if (styles != null) {
					styles[i - row][j - column] = r.getCellStyle();
				}
				if (data != null) {
					data[i - row][j - column] = ReservedCellData.reserve(r);
				}
			}
		}

		return result;
	}

	public static class ReservedResult {
		final Sheet _sheet;
		ReservedCellData[][] _data = null;
		CellStyle[][] _styles = null;
		final int _row, _column, _lastRow, _lastColumn;

		public ReservedResult(Sheet sheet, int row, int column, int lastRow,
				int lastColumn) {
			_sheet = sheet;
			_row = row;
			_column = column;
			_lastRow = lastRow;
			_lastColumn = lastColumn;
		}

		public ReservedCellData[][] getData() {
			return _data;
		}

		private void setData(ReservedCellData[][] _data) {
			this._data = _data;
		}

		public CellStyle[][] getStyles() {
			return _styles;
		}

		private void setStyles(CellStyle[][] _styles) {
			this._styles = _styles;
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

		
		public void restore(){
			int row = getRow();
			int column = getColumn();
			int lastRow = getLastRow();
			int lastColumn = getLastColumn();
			Sheet sheet = getSheet();
			
			ReservedCellData[][] data = getData();
			CellStyle[][] styles = getStyles();
			
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
	
	public static class ReservedCellData {

		boolean _blank = false;;
		String _editText;
		
		public ReservedCellData(){
			this._blank = true;
		}
		
		public ReservedCellData(String editText){
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
		
		public static ReservedCellData reserve(Range range){
			CellData d = range.getCellData();
			
			if(d.isBlank()){
				return new ReservedCellData();
			}
			
			String editText = d.getEditText();
			//TODO handle other data someday(hyperlink, comment)
			return new ReservedCellData(editText);
		}
	}

}
