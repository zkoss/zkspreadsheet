package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NDataRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCell.CellType;
/**
 * The implementation that store back value to cell
 * @author dennis
 *
 */
public class LocalValueDataGridImpl implements NDataGrid,Serializable {
	private static final long serialVersionUID = 1L;
	protected final NSheet sheet;
	protected LocalValueDataGridImpl(NSheet sheet) {
		this.sheet = sheet;
	}

	@Override
	public NCellValue getValue(int rowIdx,int columnIdx) {
		return getLocalValue(rowIdx,columnIdx);
	}
	
	protected NCellValue getLocalValue(int rowIdx,int columnIdx) {
		NCell cell = sheet.getCell(rowIdx, columnIdx);
		if(cell instanceof AbstractCellAdv){
			return ((AbstractCellAdv)cell).getLocalValue();
		}
		throw new IllegalStateException("doesn't support to get value from cell "+cell);
	}

	@Override
	public void setValue(int rowIdx,int columnIdx, NCellValue value) {
		setLocalValue(rowIdx,columnIdx,value);
	}
	
	protected void setLocalValue(int rowIdx,int columnIdx, NCellValue value) {
		NCell cell = sheet.getCell(rowIdx, columnIdx);
		if(cell instanceof AbstractCellAdv){
			((AbstractCellAdv)cell).setLocalValue((value==null||value.getType()==CellType.BLANK)?null:value);
		}else{
			throw new IllegalStateException("doesn't support to store value to cell "+cell);
		}
	}
	
	@Override
	public boolean validateValue(int rowIdx,int columnIdx, NCellValue value) {
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
	public Iterator<NDataRow> getRowIterator() {
		throw new UnsupportedOperationException("doens't support it");
	}

}
