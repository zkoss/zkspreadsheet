package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.util.CellReference;

class CellProxyImpl implements NCell{
	SheetImpl sheet;
	int row;
	int column;
	CellImpl proxy;
	
	
	public CellProxyImpl(SheetImpl sheet, int row,int column) {
		this.sheet = sheet;
		this.row = row;
		this.column = column;
	}
	
	
	private void loadProxy(){
		if(proxy==null){
			proxy = (CellImpl)sheet.getCellAt(row, column, false);
		}
	}

	public boolean isNull() {
		loadProxy();
		return proxy==null?true:proxy.isNull();
	}


	public CellType getType() {
		loadProxy();
		return proxy==null?CellType.BLANK:proxy.getType();
	}


	public int getRowIndex() {
		loadProxy();
		return proxy==null?row:proxy.getRowIndex();
	}


	public int getColumnIndex() {
		loadProxy();
		return proxy==null?column:proxy.getRowIndex();
	}


	public void setValue(Object value) {
		loadProxy();
		if(proxy==null){
			proxy = (CellImpl)((RowImpl)sheet.getOrCreateRowAt(row)).getOrCreateCellAt(column);
		}
		proxy.setValue(value);
	}


	public Object getValue() {
		loadProxy();
		return proxy==null?null:proxy.getValue();
	}

	public String asString(boolean enableSheetName) {
		loadProxy();
		return proxy==null?new CellReference(enableSheetName?sheet.getSheetName():null, row,column,false,false).formatAsString():
			proxy.asString(enableSheetName);
	}

}
