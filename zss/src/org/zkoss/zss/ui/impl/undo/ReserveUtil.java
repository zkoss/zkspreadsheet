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
package org.zkoss.zss.ui.impl.undo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.CellStyleHelper;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.CellRegion;

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
		
		AreaRef[] mergeInfo = null;
		
		int rowStart,rowEnd;
		rowStart=rowEnd=-1;
		
		Map<Integer,ReservedRow> reservedRows = null;
		
//		Range r = Ranges.range(sheet, row, column, lastRow, lastColumn);
//		if(r.isWholeRow() && r.isWholeColumn()){
//			throw new IllegalOpArgumentException("doesn't support to reserve data when select all");
//		}
		
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
				
				int first,last;
				first = sheet.getFirstColumn(i);//Math.max(_sheet.getFirstColumn(i),r.getColumn());
				last = sheet.getLastColumn(i);//Math.min(_sheet.getLastColumn(i),r.getLastColumn());//-1 if no such col
				
				if(lastColumn<first || column>last){
					//not overlap
					first = last = -1;
				}else{
					first = Math.max(first,column);
					last = Math.min(last,lastColumn);
				}
				
				ReservedRow reservedRow = new ReservedRow(/*i,*/first,last);
				reservedRows.put(i, reservedRow);
				if(first>=0 && last>=0){
					for(int j=first;j<=last;j++){

						Range range = Ranges.range(sheet,i,j);
						ReservedCell cell = new ReservedCell(/*i, j*/);
						reservedRow.setCell(j, cell);
						
						if(reserveContent){
							ReservedCellContent content = ReservedCellContent.reserve(range);
							cell.setContent(content);
						}
						
						if(reserveStyle){
							ReservedCellStyle style = ReservedCellStyle.reserve(range);
							cell.setStyle(style);
						}
					}
				}
			}
		}
		ReservedResult result = new ReservedResult(sheet,row,column,lastRow,lastColumn,reserveType);
		result.setRowsInfo(reservedRows, rowStart, rowEnd);
		
		if(reserveMerge){
			result.setMergeInfo(mergeInfo = reserveMergeInfo(sheet,row,column,lastRow,lastColumn));
		}
		return result;
	}
	
	/**
	 * reserve the merge information that in the given range.
	 */
	public static AreaRef[] reserveMergeInfo(Sheet sheet, int row, int column,
			int lastRow, int lastColumn){
		int size = ((SheetImpl)sheet).getNative().getNumOfMergedRegion();
		ArrayList<AreaRef> array = new ArrayList<AreaRef>();
		AreaRef cur = new AreaRef(row,column,lastRow,lastColumn);
		for(int i=0;i<size;i++){
			CellRegion cra = ((SheetImpl)sheet).getNative().getMergedRegion(i);
			int r = cra.row;
			int c = cra.column;
			int lr = cra.lastRow;
			int lc = cra.lastColumn;
			if(cur.contains(r, c, lr, lc)){
				array.add(new AreaRef(r,c,lr,lc));
			}
		}
		return array.size()==0?null:array.toArray(new AreaRef[array.size()]);
	}

	public static class ReservedResult {
		final int _reserveType;
		final Sheet _sheet;
		Map<Integer,ReservedRow> _rows = null;
		int _rowStart,_rowEnd;
		AreaRef[] _mergeInfo;
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
		
		public AreaRef[] getMergeInfo() {
			return _mergeInfo;
		}

		public void setMergeInfo(AreaRef[] mergeInfo) {
			this._mergeInfo = mergeInfo;
		}
		public void setRowsInfo(Map<Integer,ReservedRow> rows, int rowStart,int rowEnd){
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
			Range r;
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
			
			AreaRef[] mergeInfo = getMergeInfo(); 
			
			//clear current merge info, we will restore it later
			if(reserveMerge){
				//clear the merge first
				AreaRef[] curMergeInfo = reserveMergeInfo(_sheet,_row,_column,_lastRow,_lastColumn);
				if(curMergeInfo!=null){
					for(AreaRef rect:curMergeInfo){
						r = Ranges.range(_sheet,rect);
						r.unmerge();
					}
				}
			}
			
			//clear content, and we will restore it back later.
			//following, it just clear whole area , it waste time because of reservation area will fill back. 
			r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
			if(reserveContent){
				r.clearContents();
			}
			if(reserveStyle){
				r.clearStyles();
			}
			
			//start to restore
			if(_rowStart>=0 && _rowEnd>=0 && (reserveContent || reserveStyle)){
				for(int i=_rowStart;i<=_rowEnd;i++){
					ReservedRow reservedRow = _rows.get(i);
					int colStart = reservedRow.getColumnStart();
					int colEnd = reservedRow.getColumnEnd();
					if(colStart>=0 && colEnd>=0){
						for(int j=colStart;j<=colEnd;j++){
							ReservedCell reservedCell = reservedRow.getCell(j);
							r = Ranges.range(_sheet,i,j);
							if(reserveContent){
								ReservedCellContent data = reservedCell.getContent();
								data.apply(r);
							}
							if(reserveStyle){
								ReservedCellStyle style = reservedCell.getStyle();
								style.apply(r);
							}
						}
					}
				}
			}

			//restore merge area
			if(reserveMerge){
				//restore merge
				if(mergeInfo!=null){				
					for(AreaRef rect:mergeInfo){
						r = Ranges.range(_sheet,rect);
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
		private ReservedCellStyle _style;
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
		public ReservedCellStyle getStyle() {
			return _style;
		}
		public void setStyle(ReservedCellStyle style) {
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
	
	public static class ReservedCellStyle {

		Font _font;
		CellStyle _style;
		
		public ReservedCellStyle(CellStyle style){
			this._style = style;
			this._font = style.getFont();
		}
		
		public void apply(Range range){
			CellStyleHelper helper = range.getCellStyleHelper();
			
			
			if(helper.isAvailable(_style)){
				range.setCellStyle(_style);
			}else{
				//TODO fix ZSS-424 get exception when undo after save
			}
		}
		
		public static ReservedCellStyle reserve(Range range){
			return new ReservedCellStyle(range.getCellStyle());
		}
	}
}
