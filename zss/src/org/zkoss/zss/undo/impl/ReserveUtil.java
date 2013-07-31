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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.IllegalOpArgumentException;
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
	public static final int RESERVE_CONTENT = 1;
	public static final int RESERVE_STYLE = 2;
	public static final int RESERVE_MERGE = 4;
	public static final int RESERVE_ALL = RESERVE_CONTENT|RESERVE_STYLE|RESERVE_MERGE;
	public static ReservedResult reserve(Sheet sheet, int row, int column,
			int lastRow, int lastColumn, int reserveType) {

		boolean reserveContent = false;
		boolean reserveStyle = false;
		boolean reserveMerge = false;
		if((reserveType & RESERVE_CONTENT)!=0){
			reserveContent = true;
		}
		if((reserveType & RESERVE_STYLE)!=0){
			reserveStyle = true;
		}
		if((reserveType & RESERVE_MERGE)!=0){
			reserveMerge = true;
		}
		
		Rect[] mergeInfo = null;
		
		int rowStart=-1;
		int rowEnd=-1;
		Map<Integer,ReservedRow> reservedRows = null;
		
		Range r = Ranges.range(sheet, row, column, lastRow, lastColumn);
		if(r.isWholeRow() && r.isWholeColumn()){
			throw new IllegalOpArgumentException("doesn't support to do this when select all");
		}
		
		rowStart = sheet.getFirstRow();
		rowEnd = sheet.getLastRow();
		if(lastRow<rowStart || row>rowEnd){
			//not overlap, reserve nothing
			rowStart = rowEnd = -1;
		}else{
			rowStart = Math.max(rowStart,row);
			rowEnd = Math.min(rowEnd,lastRow);
		}
		
		if(rowStart>=0 && rowEnd>=0 && (reserveContent || reserveStyle)){
			reservedRows = new HashMap<Integer, ReservedRow>(rowEnd-rowStart+1);
			for(int i=rowStart;i<=rowEnd;i++){
				
				int colStart,colEnd;
				colStart = sheet.getFirstColumn(i);//Math.max(_sheet.getFirstColumn(i),r.getColumn());
				colEnd = sheet.getLastColumn(i);//Math.min(_sheet.getLastColumn(i),r.getLastColumn());//-1 if no such col
				
				if(lastColumn<colStart || column>colEnd){
					//not overlap
					colStart = colEnd = -1;
				}else{
					colStart = Math.max(colStart,column);
					colEnd = Math.min(colEnd,lastColumn);
				}
				
				ReservedRow reservedRow = new ReservedRow(/*i,*/colStart,colEnd);
				reservedRows.put(i, reservedRow);
				if(colStart>=0 && colEnd>=0){
					for(int j=colStart;j<=colEnd;j++){

						Range range = Ranges.range(sheet,i,j);
						ReservedCell cell = new ReservedCell(/*i, j*/);
						reservedRow.setCell(j, cell);
						
						if(reserveContent){
							ReservedCellContent content = ReservedCellContent.reserve(range);
							cell.setContent(content);
						}
						
						if(reserveStyle){
							CellStyle style = range.getCellStyle();
							
							cell.setStyle(style);
						}
					}
				}
			}
		}
		ReservedResult result = new ReservedResult(sheet,row,column,lastRow,lastColumn,reserveType);
		result.setRowInfo(reservedRows, rowStart, rowEnd);
		
		if(reserveMerge){
			result.setMergeInfo(mergeInfo = reserveMergeInfo(sheet,row,column,lastRow,lastColumn));
		}
		return result;
	}
	
	/**
	 * reserve the merge information that in the given range.
	 */
	public static Rect[] reserveMergeInfo(Sheet sheet, int row, int column,
			int lastRow, int lastColumn){
		org.zkoss.poi.ss.usermodel.Sheet poiSheet = sheet.getPoiSheet();
		int size = poiSheet.getNumMergedRegions();
		ArrayList<Rect> array = new ArrayList<Rect>();
		Rect cur = new Rect(column,row,lastColumn,lastRow);
		for(int i=0;i<size;i++){
			CellRangeAddress cra = poiSheet.getMergedRegion(i);
			int r = cra.getFirstRow();
			int c = cra.getFirstColumn();
			int lr = cra.getLastRow();
			int lc = cra.getLastColumn();
			if(cur.contains(r, c, lr, lc)){
				array.add(new Rect(c,r,lc,lr));
			}
		}
		return array.size()==0?null:array.toArray(new Rect[array.size()]);
	}

	public static class ReservedResult {
		final int _reserveType;
		final Sheet _sheet;
		Map<Integer,ReservedRow> _rows = null;
		int _rowStart,_rowEnd;
		Rect[] _mergeInfo;
		final int _row, _column, _lastRow, _lastColumn;

		public ReservedResult(Sheet sheet, int row, int column, int lastRow,
				int lastColumn, int reserveType) {
			_sheet = sheet;
			_row = row;
			_column = column;
			_lastRow = lastRow;
			_lastColumn = lastColumn;
			_reserveType = reserveType;
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
		public void setRowInfo(Map<Integer,ReservedRow> rows, int rowStart, int rowEnd){
			_rows = rows;
			_rowStart = rowStart;
			_rowEnd = rowEnd;
		}
		
		public Map<Integer,ReservedRow> getRows(){
			return _rows;
		}
		public int getRowStart(){
			return _rowStart;
		}
		public int getColumnEnd(){
			return _rowEnd;
		}

		public void restore(){
			int row = getRow();
			int column = getColumn();
			int lastRow = getLastRow();
			int lastColumn = getLastColumn();
			Sheet sheet = getSheet();
			boolean reserveContent = false;
			boolean reserveStyle = false;
			boolean reserveMerge = false;
			if((_reserveType & RESERVE_CONTENT)!=0){
				reserveContent = true;
			}
			if((_reserveType & RESERVE_STYLE)!=0){
				reserveStyle = true;
			}
			if((_reserveType & RESERVE_MERGE)!=0){
				reserveMerge = true;
			}
			
			Rect[] mergeInfo = getMergeInfo(); 
			
			//
			if(reserveMerge){
				//clear the merge first
				Rect[] curMergeInfo = reserveMergeInfo(sheet,row,column,lastRow,lastColumn);
				if(curMergeInfo!=null){
					for(Rect rect:curMergeInfo){
						Range r = Ranges.range(sheet,rect);
						r.unmerge();
					}
				}
			}
			
			//clear content before store back
			Range r = Ranges.range(sheet,row,column,lastRow,lastColumn);
			if(reserveContent){
				r.clearContents();
			}
			if(reserveStyle){
				r.clearStyles();
			}
			
			if(_rowStart>=0 && _rowEnd>=0 && (reserveContent || reserveStyle)){
				for(int i=_rowStart;i<=_rowEnd;i++){
					ReservedRow reservedRow = _rows.get(i);
					int colStart = reservedRow.getColumnStart();
					int colEnd = reservedRow.getColumnEnd();
					if(colStart>=0 && colEnd>=0){
						for(int j=colStart;j<=colEnd;j++){
							ReservedCell reservedCell = reservedRow.getCell(j);
							Range range = Ranges.range(_sheet,i,j);
							if(reserveContent){
								ReservedCellContent data = reservedCell.getContent();
								data.apply(range);
							}
							if(reserveStyle){
								CellStyle style = reservedCell.getStyle();
								range.setCellStyle(style);
							}
						}
					}
				}
				
				//clear content not in reserved range
			}

			
			
			
			if(reserveMerge){
				//restore merge
				if(mergeInfo!=null){				
					for(Rect rect:mergeInfo){
						r = Ranges.range(sheet,rect);
						r.merge(false);
					}
				}
			}

		}
	}
	
	public static class ReservedRow {
//		private int _index;
		private int _colStart;
		private int _colEnd;
		private Map<Integer,ReservedCell> cells;
		public ReservedRow(/*int index,*/int colStart,int colEnd){
//			this._index = index;
			this._colStart = colStart;
			this._colEnd = colEnd;
		}
		
		public void setCell(int col, ReservedCell cell){
			if(col<_colStart&&col>_colEnd){
				throw new IllegalArgumentException("not in range "+_colStart+","+_colEnd);
			}
			if(cells==null){
				cells = new HashMap<Integer, ReservedCell>();
			}
			cells.put(col, cell);
		}
		
		public ReservedCell getCell(int col){
			return cells==null?null:cells.get(col);
		}
		
		public int getColumnStart(){
			return _colStart;
		}
		
		public int getColumnEnd(){
			return _colEnd;
		}
	}
	
	public static class ReservedCell {
//		private int _row;
//		private int _column;
		private CellStyle _style;
		private ReservedCellContent _content;
		public ReservedCell(/*int row,int column*/){
//			this._row = row;
//			this._column = column;
		}
		public ReservedCellContent getContent() {
			return _content;
		}
		public void setContent(ReservedCellContent content) {
			this._content = content;
		}
		public CellStyle getStyle() {
			return _style;
		}
		public void setStyle(CellStyle style) {
			this._style = style;
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
