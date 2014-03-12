/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.range.impl;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
public class ModelUpdate {
	public static enum UpdateType{
		REF, REFS, CELL, CELLS, MERGE, INSERT_DELETE
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

	@Override
	public String toString() {
		return "ModelUpdate [type=" + type + ", data=" + data + "]";
	}
	
	
}
