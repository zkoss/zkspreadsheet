package org.zkoss.zss.ngmodel.input;

import org.zkoss.zss.ngmodel.NCell.CellType;

public class InputResult {

	private String input = null;
	private Object value = null;
	private CellType type = CellType.BLANK;
	public InputResult(){}
	public InputResult(String input) {
		this.input = input;
		this.type = type;
		this.value = value;
	}

	public String getInput() {
		return input;
	}

	public Object getValue() {
		return value;
	}

	public CellType getType() {
		return type;
	}

	void setValue(Object value) {
		this.value = value;
	}

	void setType(CellType type) {
		this.type = type;
	}

	
}
