/* DeleteCellAction.java

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

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.IllegalOpArgumentException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.impl.undo.ReserveUtil.ReservedCellContent;
/**
 * 
 * @author dennis
 *
 */
public class DeleteCellAction extends AbstractUndoableAction {
	
	DeleteShift _shift;
	int _rowStart=-1;
	int _rowEnd=-1;
	Map<Integer,ReservedRow> _rows;
	Map<Integer,Integer> _rowHeights;
	Map<Integer,Integer> _colWidths;
	AreaRef[] _mergeInfo;
	boolean _doFlag;
	
	
	public DeleteCellAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,DeleteShift shift){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._shift = shift;
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		_doFlag = true;
		Range r = Ranges.range(_sheet, _row, _column, _lastRow, _lastColumn);
		if(r.isWholeRow() && r.isWholeColumn()){
			throw new IllegalOpArgumentException("doesn't support to delete all");
		}else if(r.isWholeColumn()){
			_colWidths = new HashMap<Integer,Integer>();
		}else if(r.isWholeRow()){
			_rowHeights = new HashMap<Integer,Integer>();
		}
		
		_rowStart = _sheet.getFirstRow();
		_rowEnd = _sheet.getLastRow();
		if(_lastRow<_rowStart || _row>_rowEnd){
			//not overlap
			_rowStart = _rowEnd = -1;
		}else{
			_rowStart = Math.max(_rowStart,_row);
			_rowEnd = Math.min(_rowEnd,_lastRow);
		}
		
		if(_rowStart>=0 && _rowEnd>=0){
			_rows = new HashMap<Integer, DeleteCellAction.ReservedRow>(_rowEnd-_rowStart+1);
			for(int i=_rowStart;i<=_rowEnd;i++){
				
				int colStart,colEnd;
				colStart = _sheet.getFirstColumn(i);//Math.max(_sheet.getFirstColumn(i),r.getColumn());
				colEnd = _sheet.getLastColumn(i);//Math.min(_sheet.getLastColumn(i),r.getLastColumn());//-1 if no such col
				
				if(_lastColumn<colStart || _column>colEnd){
					//not overlap
					colStart = colEnd = -1;
				}else{
					colStart = Math.max(colStart,_column);
					colEnd = Math.min(colEnd,_lastColumn);
				}
				
				if(_rowHeights!=null){
					_rowHeights.put(i, _sheet.getRowHeight(i));
				}
				
				ReservedRow row = new ReservedRow(/*i,*/colStart,colEnd);
				_rows.put(i, row);
				if(colStart>=0 && colEnd>=0){
					for(int j=colStart;j<=colEnd;j++){
						if(i==_rowStart && _colWidths!=null){
							_colWidths.put(j, _sheet.getColumnWidth(i));
						}
						Range range = Ranges.range(_sheet,i,j);
						CellStyle style = range.getCellStyle();
						ReservedCellContent data = ReservedCellContent.reserve(range);
						ReservedCell cell = new ReservedCell(/*i, j*/);
						cell.setStyle(style);
						cell.setData(data);
						row.setCell(j, cell);
					}
				}
			}
		}
		//keep merge info
		_mergeInfo = ReserveUtil.reserveMergeInfo(_sheet, _row, _column, _lastRow, _lastColumn);

		CellOperationUtil.delete(r,_shift);
	}

	@Override
	public boolean isUndoable() {
		return _doFlag && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return !_doFlag && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		Range r = Ranges.range(_sheet, _row, _column, _lastRow, _lastColumn);
		switch(_shift){
			case UP:
				CellOperationUtil.insert(r,InsertShift.DOWN,InsertCopyOrigin.FORMAT_NONE);
				break;
			case LEFT:
				CellOperationUtil.insert(r,InsertShift.RIGHT,InsertCopyOrigin.FORMAT_NONE);
				break;
			case DEFAULT:
				CellOperationUtil.insert(r,InsertShift.DEFAULT,InsertCopyOrigin.FORMAT_NONE);
				break;
		}
		
		if(_mergeInfo!=null && _mergeInfo.length>0){
			for(AreaRef rect:_mergeInfo){
				//restore back
				Range mergeRange = Ranges.range(_sheet,rect);
				mergeRange.merge(false);
			}
		}
		
		if(_rowStart>=0 && _rowEnd>=0){
			for(int i=_rowStart;i<=_rowEnd;i++){
				ReservedRow row = _rows.get(i);
				int colStart = row.getColumnStart();
				int colEnd = row.getColumnEnd();
				if(_rowHeights!=null){
					Integer h = _rowHeights.get(i);
					if(h!=null){
						 Ranges.range(_sheet,i,0).setRowHeight(h);
					}
				}
				if(colStart>=0 && colEnd>=0){
					for(int j=colStart;j<=colEnd;j++){
						if(i==_rowStart && _colWidths!=null){
							Integer w = _colWidths.get(i);
							if(w!=null){
								 Ranges.range(_sheet,0,j).setColumnWidth(w);
							}
						}
						
						ReservedCell rcell = row.getCell(j);
						Range range = Ranges.range(_sheet,i,j);
						CellStyle style = rcell.getStyle();
						ReservedCellContent data = rcell.getData();
						
						range.setCellStyle(style);
						data.apply(range);
					}
				}
			}
		}
		
		
		_doFlag = false;
	}
	
	private static class ReservedRow {
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
				cells = new HashMap<Integer, DeleteCellAction.ReservedCell>();
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
	
	private static class ReservedCell {
//		private int _row;
//		private int _column;
		private CellStyle _style;
		private ReservedCellContent _data;
		public ReservedCell(/*int row,int column*/){
//			this._row = row;
//			this._column = column;
		}
		public ReservedCellContent getData() {
			return _data;
		}
		public void setData(ReservedCellContent data) {
			this._data = data;
		}
		public CellStyle getStyle() {
			return _style;
		}
		public void setStyle(CellStyle style) {
			this._style = style;
		}
		
	}
}
