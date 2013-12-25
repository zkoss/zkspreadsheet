package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NDataRow;
import org.zkoss.zss.ngmodel.NSheet;

public class DataGridImpl implements NDataGrid,Serializable {
	private static final long serialVersionUID = 1L;
	private NSheet sheet;
	public DataGridImpl(NSheet sheet) {
		this.sheet = sheet;
	}

	@Override
	public Object getValue(int rowIdx,int columnIdx) {
		NCell cell = sheet.getCell(rowIdx, columnIdx);
		if(cell instanceof AbstractCellAdv){
			return ((AbstractCellAdv)cell).getLocalValue();
		}
		throw new IllegalStateException("doesn't support to store value to cell "+cell);
	}

	@Override
	public void setValue(int rowIdx,int columnIdx, Object value) {
		NCell cell = sheet.getCell(rowIdx, columnIdx);
		if(cell instanceof AbstractCellAdv){
			((AbstractCellAdv)cell).setLocalValue(value);
		}else{
			throw new IllegalStateException("doesn't support to store value to cell "+cell);
		}
	}
	
	@Override
	public boolean validateValue(int rowIdx,int columnIdx, Object value) {
		return true;
	}


	@Override
	public boolean supportOperations() {
		return true;
	}

	@Override
	public void insertRow(int rowIdx, int size) {
		//don't need to do anything, we store data in cell local value
	}

	@Override
	public void deleteRow(int rowIdx, int size) {
		//don't need to do anything, we store data in cell local value
	}

	@Override
	public void insertColumn(int rowIdx, int size) {
		//don't need to do anything, we store data in cell local value
	}

	@Override
	public void deleteColumn(int rowIdx, int size) {
		//don't need to do anything, we store data in cell local value
	}

	@Override
	public boolean supportDataIterator() {
		return false;
	}

	@Override
	public Iterator<NDataRow> getDataRowIterator() {
		//not support
		return null;
	}

}
