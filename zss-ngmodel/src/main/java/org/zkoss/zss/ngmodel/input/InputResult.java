package org.zkoss.zss.ngmodel.input;

import org.zkoss.zss.ngmodel.NCell.CellType;

public class InputResult {

	CellType type = CellType.BLANK;
	
	Object value = null;

	public CellType getType() {
		return type;
	}

	public void setType(CellType type) {
		this.type = type;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}
}
