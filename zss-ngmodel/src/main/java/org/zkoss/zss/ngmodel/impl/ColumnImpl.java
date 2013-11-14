package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NColumn;

public class ColumnImpl implements NColumn{

	SheetImpl sheet;
	
	public ColumnImpl(SheetImpl sheet) {
		this.sheet = sheet;
	}

	public int getIndex() {
		return sheet.getColumnIndex(this);
	}

	public boolean isNull() {
		return false;
	}

	protected void onModelEvent(ModelEvent event) {
		// TODO Auto-generated method stub
		
	}

}
