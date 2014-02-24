package org.zkoss.zss.ngmodel.impl.sys;

import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.input.InputResult;

public class InputResultImpl implements InputResult{

	private String editText = null;
	private Object value = null;
	private CellType type = CellType.BLANK;
	private String format = null;
	public InputResultImpl(){}
	public InputResultImpl(String input) {
		this.editText = input;
	}

	public String getEditText() {
		return editText;
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
	public String getFormat() {
		return format;
	}
	void setFormat(String format) {
		this.format = format;
	}
}
