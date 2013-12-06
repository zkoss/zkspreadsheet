/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.util.HashMap;
import java.util.Map;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ModelEvent {

	private String name;
	
	private Map<String,Object> data;
	
	public ModelEvent(String name){
		this.name = name;
	}
	
	public ModelEvent(String name,Map<String,Object> data){
		this.name = name;
		this.data = new HashMap<String, Object>(data);
	}
	
	public Object getData(String key){
		return data==null?null:data.get(key);
	}

	public String getName() {
		return name;
	}
	
}
