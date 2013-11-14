package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;

public class CellImpl implements NCell {

	RowImpl row;
	CellType type = CellType.BLANK;
	Object value = null;
	
	public CellImpl(RowImpl row){
		this.row = row;
	}
	
	public CellType getType() {
		return type;
	}

	public boolean isNull() {
		return false;
	}

	public int getRowIndex() {
		return row.getIndex();
	}

	public int getColumnIndex() {
		return row.getCellIndex(this);
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}

}
