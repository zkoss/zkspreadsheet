package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.util.CellReference;

class ColumnProxyImpl implements NColumn{
	SheetImpl sheet;
	int index;
	ColumnImpl proxy;
	
	
	public ColumnProxyImpl(SheetImpl sheet, int index) {
		this.sheet = sheet;
		this.index = index;
	}
	
	
	private void loadProxy(){
		if(proxy==null){
			proxy = (ColumnImpl)sheet.getColumnAt(index,false);
		}
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

}
