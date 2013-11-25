package org.zkoss.zss.ngmodel;

import java.util.HashMap;
import java.util.Map;

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
