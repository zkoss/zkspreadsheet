package org.zkoss.zss.range.impl;

public class ModelUpdate {
	public static enum UpdateType{
		REF, REFS, CELL, MERGE, INSERT_DELETE
	}
	
	final UpdateType type;
	final Object data;
	
	public ModelUpdate(UpdateType type,Object data){
		this.type = type;
		this.data = data;
	}
	
	public UpdateType getType(){
		return type;
	}
	
	public Object getData(){
		return data;
	}
}
