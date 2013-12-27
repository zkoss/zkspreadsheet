package org.zkoss.zss.ngmodel;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NCell.CellType;

public class NCellValue implements Serializable {
	private static final long serialVersionUID = 1L;
	private final CellType type;
	private final Object value;
	public NCellValue(String value){
		this(CellType.STRING,value);
	}
	public NCellValue(Double number){
		this(CellType.NUMBER,number);
	}
	public NCellValue(Boolean bool){
		this(CellType.BOOLEAN,bool);
	}
	public NCellValue(){
		this(CellType.BLANK,null);
	}
	
	protected NCellValue(CellType type, Object value){
		this.type = value==null?CellType.BLANK:type;
		this.value = value;
	}
	public CellType getType() {
		return type;
	}
	public Object getValue() {
		return value;
	}
}
