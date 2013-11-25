package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.ModelEventListener;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;

public abstract class AbstractBook implements NBook,Serializable{
	private static final long serialVersionUID = 1L;

	List<ModelEventListener> listeners;
	
	/**
	 * Optimize CellStyle, usually called when export book. 
	 * @return
	 */
	/*package*/ abstract List<NCell> optimizeCellStyle();
	
	
	protected void sendEvent(String name, Object... data){
		Map<String,Object> datamap = new HashMap<String,Object>();
		datamap.put("book", this);
		if(datamap!=null){
			if(data.length%2 != 0){
				throw new IllegalArgumentException("event data must be key,value pair");
			}
			for(int i=0;i<data.length;i+=2){
				if(!(data[i] instanceof String)){
					throw new IllegalArgumentException("event data key must be string");
				}
				datamap.put((String)data[i],data[i+1]);
			}
		}
		ModelEvent event = new ModelEvent(name, datamap);
		sendEvent(event);
	}
	
	protected void sendEvent(ModelEvent event){
		if(listeners ==null)
			return;
		for(ModelEventListener l:listeners){
			l.onEvent(event);
		}
	}
	
	public void addEventListener(ModelEventListener listener){
		if(listeners==null){
			listeners = new LinkedList<ModelEventListener>();
		}
		if(!listeners.contains(listener)){
			listeners.add(listener);
		}
	}
	
	public void removeEventListener(ModelEventListener listener){
		if(listeners!=null){
			listeners.remove(listener);
		}
	}
	
}
