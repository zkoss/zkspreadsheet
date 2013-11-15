package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.util.CellReference;

class ColumnProxyImpl implements NColumn{
	WeakReference<SheetImpl> sheetRef;
	int index;
	ColumnImpl proxy;
	
	
	public ColumnProxyImpl(SheetImpl sheet, int index) {
		this.sheetRef = new WeakReference(sheet);
		this.index = index;
	}
	
	
	protected void loadProxy(){
		if(proxy==null){
			proxy = (ColumnImpl)getSheet().getColumnAt(index,false);
			if(proxy!=null){
				sheetRef.clear();
			}
		}
	}
	
	protected SheetImpl getSheet(){
		SheetImpl sheet = sheetRef.get();
		if(sheet==null){
			throw new IllegalStateException("proxy target lost, you should't keep this instance");
		}
		return sheet;
	}
	
	public int getIndex() {
		loadProxy();
		return proxy==null?index:proxy.getIndex();
	}


	public boolean isNull() {
		loadProxy();
		return proxy==null?true:proxy.isNull();
	}
	
	public String asString() {
		loadProxy();
		return proxy==null?CellReference.convertNumToColString(getIndex()):proxy.asString();
	}


	public NCellStyle getCellStyle() {
		loadProxy();
		return proxy==null?null:proxy.getCellStyle();
	}

}
