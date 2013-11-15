package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.util.CellReference;

class RowProxyImpl implements NRow{
	SheetImpl sheet;
	int index;
	RowImpl proxy;
	
	
	public RowProxyImpl(SheetImpl sheet, int index) {
		this.sheet = sheet;
		this.index = index;
	}
	
	
	private void loadProxy(){
		if(proxy==null){
			proxy = (RowImpl)sheet.getRowAt(index,false);
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


	public int getStartCellIndex() {
		loadProxy();
		return proxy==null?-1:proxy.getStartCellIndex();
	}


	public int getEndCellIndex() {
		loadProxy();
		return proxy==null?-1:proxy.getEndCellIndex();
	}

	public String asString() {
		loadProxy();
		return proxy==null?Integer.toString(index+1):proxy.asString();
	}
	
	public NCellStyle getCellStyle() {
		loadProxy();
		return proxy==null?null:proxy.getCellStyle();
	}
}
