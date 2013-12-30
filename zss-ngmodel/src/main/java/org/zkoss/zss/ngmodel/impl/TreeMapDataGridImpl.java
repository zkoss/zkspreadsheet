package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NDataCell;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NDataRow;
/**
 * The implementation that store value in tree map, this implementation is for data grid capability test
 * @author dennis
 *
 */
public class TreeMapDataGridImpl implements NDataGrid,Serializable {
	private static final long serialVersionUID = 1L;

	
	private final IndexPool<DataRowImpl> rows = new IndexPool<DataRowImpl>(){
		private static final long serialVersionUID = 1L;

		@Override
		void resetIndex(int newidx, DataRowImpl obj) {
			obj.setIndex(newidx);
		}};
	
	public TreeMapDataGridImpl() {
	}

	@Override
	public NCellValue getValue(int rowIdx, int columnIdx) {
		DataRowImpl row = getRow(rowIdx, false);
		if(row==null)
			return null;
		DataCellImpl cell = row.getCell(columnIdx, false);
		return cell==null?null:cell.getValue();
	}

	@Override
	public void setValue(int rowIdx, int columnIdx, NCellValue value) {
		DataRowImpl row = getRow(rowIdx, value!=null);
		if(row==null) return;
		DataCellImpl cell = row.getCell(columnIdx, value!=null);
		if(cell!=null){
			cell.setValue(value);
		}
	}

	@Override
	public boolean isSupportedOperations() {
		return true;
	}

	@Override
	public void insertRow(int rowIdx, int size) {
		if(size<=0) return;
		
		rows.insert(rowIdx, size);
		
//		shiftAfterRowInsert(rowIdx,size);
	}
	
	

	@Override
	public void deleteRow(int rowIdx, int size) {
		if(size<=0) return;
		
		//clear before move relation
		for(DataRowImpl row:rows.subValues(rowIdx,rowIdx+size)){
			row.destroy();
		}		
		rows.delete(rowIdx, size);
		
//		shiftAfterRowDelete(rowIdx,size);
	}

	@Override
	public void insertColumn(int columnIdx, int size) {
		if(size<=0) return;
		
		for(DataRowImpl row:rows.values()){
			row.insertCell(columnIdx,size);
		}
		
//		shiftAfterColumnInsert(columnIdx,size);
	}

	@Override
	public void deleteColumn(int columnIdx, int size) {
		if(size<=0) return;
		
		for(DataRowImpl row:rows.values()){
			row.deleteCell(columnIdx,size);
		}
//		shiftAfterColumnDelete(columnIdx,size);
	}

	@Override
	public boolean isProvidedIterator() {
		return true;
	}

	@Override
	public Iterator<NDataRow> getRowIterator() {
		return Collections.unmodifiableCollection((Collection)rows.values()).iterator();
	}

	@Override
	public boolean validateValue(int row, int column, NCellValue value) {
		return true;
	}
	
	DataRowImpl getRow(int rowIdx, boolean create){
		DataRowImpl rowObj = rows.get(rowIdx);
		if(rowObj == null && create){
			rowObj = new DataRowImpl(this,rowIdx);
			rows.put(rowIdx, rowObj);
		}
		return rowObj;
	}

	
	static class DataRowImpl implements NDataRow, Serializable{
		TreeMapDataGridImpl dataGrid;
		int rowIndex;
		private final IndexPool<DataCellImpl> cells = new IndexPool<DataCellImpl>(){
			private static final long serialVersionUID = 1L;

			@Override
			void resetIndex(int newidx, DataCellImpl obj) {
				obj.setIndex(newidx);
			}};
		
		public DataRowImpl(TreeMapDataGridImpl dataGrid,int rowIdx){
			this.dataGrid = dataGrid;
			rowIndex = rowIdx;
		}
		
		public void setIndex(int newidx) {
			this.rowIndex = newidx;
		}

		public void checkOrphan() {
			if (dataGrid == null) {
				throw new IllegalStateException("doesn't connect to parent");
			}
		}
		public void destroy() {
			checkOrphan();
			dataGrid = null;
		}
		
		@SuppressWarnings("unchecked")
		@Override
		public Iterator<NDataCell> getCellIterator() {
			return Collections.unmodifiableCollection((Collection)cells.values()).iterator();
		}

		DataCellImpl getCell(int columnIdx, boolean create){
			DataCellImpl cellObj = cells.get(columnIdx);
			if (cellObj == null && create) {
				checkOrphan();
				cellObj = new DataCellImpl(this, columnIdx);
				cells.put(columnIdx, cellObj);
			}
			return cellObj;
		}
		
		
		public void insertCell(int cellIdx, int size) {
			if (size <= 0)
				return;

			cells.insert(cellIdx, size);
		}

		public void deleteCell(int cellIdx, int size) {
			if (size <= 0)
				return;
			// clear before move relation
			for (DataCellImpl cell : cells.subValues(cellIdx, cellIdx + size)) {
				cell.destroy();
			}
			cells.delete(cellIdx, size);
		}

		@Override
		public int getIndex() {
			return rowIndex;
		}

		public int getStartCellIndex() {
			return cells.firstKey();
		}

		public int getEndCellIndex() {
			return cells.lastKey();
		}
	}
	
	static class DataCellImpl implements NDataCell, Serializable{
		private static final long serialVersionUID = 1L;
		DataRowImpl row;
		int columnIndex;
		NCellValue value;
		
		public DataCellImpl(DataRowImpl row,int columnIdx){
			this.row = row;
			this.columnIndex = columnIdx;
		}
		
		public void setValue(NCellValue value) {
			this.value = value;
		}

		public void destroy() {
			checkOrphan();
			row = null;
		}

		public void setIndex(int newidx) {
			this.columnIndex = newidx;
		}

		@Override
		public NCellValue getValue() {
			return value;
		}

		@Override
		public int getRowIndex() {
			checkOrphan();
			return row.getIndex();
		}

		@Override
		public int getColumnIndex() {
			return columnIndex;
		}
		
		public void checkOrphan() {
			if (row == null) {
				throw new IllegalStateException("doesn't connect to parent");
			}
		}
	}

	@SuppressWarnings("unchecked")
	@Override
	public Iterator<NDataCell> getCellIterator(int rowIdx) {
		DataRowImpl row = getRow(rowIdx,false);
		if(row==null){
			return Collections.EMPTY_LIST.iterator();
		}
		return row.getCellIterator();
	}

	public int getStartRowIndex() {
		return rows.firstKey();
	}

	public int getEndRowIndex() {
		return rows.lastKey();
	}

	public int getStartCellIndex(int row) {
		DataRowImpl rowObj = (DataRowImpl) getRow(row,false);
		if(rowObj!=null){
			return rowObj.getStartCellIndex();
		}
		return -1;
	}

	public int getEndCellIndex(int row) {
		DataRowImpl rowObj = (DataRowImpl) getRow(row,false);
		if(rowObj!=null){
			return rowObj.getEndCellIndex();
		}
		return -1;
	}

	@Override
	public boolean isProvidedStartEndIndex() {
		return true;
	}

}
