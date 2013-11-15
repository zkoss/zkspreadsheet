package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.util.CellReference;

class CellProxyImpl implements NCell{
	WeakReference<SheetImpl> sheetRef;
	int row;
	int column;
	CellImpl proxy;
	
	
	public CellProxyImpl(SheetImpl sheet, int row,int column) {
		this.sheetRef = new WeakReference<SheetImpl>(sheet);;
		this.row = row;
		this.column = column;
	}
	
	private SheetImpl getSheet(){
		SheetImpl sheet = sheetRef.get();
		if(sheet==null){
			throw new IllegalStateException("proxy target lost, you should't keep this instance");
		}
		return sheet;
	}
	
	private void loadProxy(){
		if(proxy==null){
			proxy = (CellImpl)getSheet().getCellAt(row, column, false);
			if(proxy!=null){
				sheetRef.clear();
			}
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
			proxy = (CellImpl)((RowImpl)getSheet().getOrCreateRowAt(row)).getOrCreateCellAt(column);
		}
		proxy.setValue(value);
	}


	public Object getValue() {
		loadProxy();
		return proxy==null?null:proxy.getValue();
	}

	public String asString(boolean enableSheetName) {
		loadProxy();
		return proxy==null?new CellReference(enableSheetName?getSheet().getSheetName():null, row,column,false,false).formatAsString():
			proxy.asString(enableSheetName);
	}

}
