package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NDataGrid;

public class DataGridImpl implements NDataGrid,Serializable {
	
	public DataGridImpl() {
	}

	@Override
	public Object getValue(NCell cell) {
		if(cell instanceof AbstractCellAdv){
			return ((AbstractCellAdv)cell).getLocalValue();
		}
		throw new IllegalStateException("doesn't allow to store value to cell "+cell);
	}

	@Override
	public void setValue(NCell cell, Object value) {
		if(cell instanceof AbstractCellAdv){
			((AbstractCellAdv)cell).setLocalValue(value);
		}else{
			throw new IllegalStateException("doesn't allow to store value to cell "+cell);
		}
	}

	@Override
	public boolean valideValue(NCell cell, Object value) {
		return true;
	}

	@Override
	public boolean supportInsertDelete() {
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

}
