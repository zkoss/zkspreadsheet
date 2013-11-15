package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.util.CellReference;

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
		checkOrphan();
		return row.getIndex();
	}

	public int getColumnIndex() {
		checkOrphan();
		return row.getCellIndex(this);
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}

	public String asString(boolean enableSheetName) {
		return new CellReference(enableSheetName?row.sheet.getSheetName():null, this).formatAsString();
	}
	
	
	protected void checkOrphan(){
		if(row==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	protected void release(){
		row = null;
	}

}
