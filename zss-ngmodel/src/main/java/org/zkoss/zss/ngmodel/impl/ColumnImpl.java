package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.util.CellReference;

public class ColumnImpl implements NColumn{

	SheetImpl sheet;
	CellStyleImpl cellStyle;
	
	public ColumnImpl(SheetImpl sheet) {
		this.sheet = sheet;
	}

	public int getIndex() {
		checkOrphan();
		return sheet.getColumnIndex(this);
	}

	public boolean isNull() {
		return false;
	}

	protected void onModelEvent(ModelEvent event) {
		// TODO Auto-generated method stub
		
	}

	public String asString() {
		return CellReference.convertNumToColString(getIndex());
	}
	
	protected void checkOrphan(){
		if(sheet==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	protected void release(){
		sheet = null;
	}

	public NCellStyle getCellStyle() {
		return cellStyle;
	}

}
