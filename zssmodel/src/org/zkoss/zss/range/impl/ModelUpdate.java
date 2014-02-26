package org.zkoss.zss.range.impl;

public class ModelUpdate {
	public static enum UpdateType{
		REF, REFS, CELL, MERGE, INSERT_DELETE
	}
	
	UpdateType type;
	Object data;
	
	public ModelUpdate(UpdateType type,Object data){
		this.data = data;
	}
	
	public UpdateType getType(){
		return type;
	}
	
	public Object getData(){
		return data;
	}
	
	public void setData(Object data){
		this.data = data;
	}
}
