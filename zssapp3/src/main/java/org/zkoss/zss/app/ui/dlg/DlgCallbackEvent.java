package org.zkoss.zss.app.ui.dlg;

import java.util.Map;

import org.zkoss.zk.ui.event.Event;

public class DlgCallbackEvent extends Event{

	public DlgCallbackEvent(String name, Map<String,Object> data) {
		super(name, null, data);
	}
	
	public Map<String,Object> getData(){
		return (Map<String,Object>)super.getData();
	}
	
	public Object getData(String name){
		Map<String,Object> data = getData();
		if(data!=null){
			return data.get(name);
		}
		return null;
	}

}
