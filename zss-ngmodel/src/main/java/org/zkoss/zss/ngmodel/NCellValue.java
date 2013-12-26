package org.zkoss.zss.ngmodel;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NCell.CellType;

public class NCellValue implements Serializable {
	private final CellType type;
	private final Object value;
	public NCellValue(CellType type, Object value){
		this.type = type;
		this.value = value;
	}
	public CellType getType() {
		return type;
	}
	public Object getValue() {
		return value;
	}
}
