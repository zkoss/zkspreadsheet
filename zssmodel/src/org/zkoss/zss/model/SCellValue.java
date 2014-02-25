package org.zkoss.zss.model;

import java.io.Serializable;

import org.zkoss.zss.model.SCell.CellType;

public class SCellValue implements Serializable {
	private static final long serialVersionUID = 1L;
	protected CellType cellType;
	protected Object value;
	public SCellValue(String value){
		this(CellType.STRING,value);
	}
	public SCellValue(Double number){
		this(CellType.NUMBER,number);
	}
	public SCellValue(Boolean bool){
		this(CellType.BOOLEAN,bool);
	}
	public SCellValue(){
		this(CellType.BLANK,null);
	}
	
	protected SCellValue(CellType type, Object value){
		this.cellType = value==null?CellType.BLANK:type;
		this.value = value;
	}
	
	public CellType getType() {
		return cellType;
	}
	public Object getValue() {
		return value;
	}
}
