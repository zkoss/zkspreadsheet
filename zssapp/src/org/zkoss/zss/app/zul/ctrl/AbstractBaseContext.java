/* AbstractBaseContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 6:22:52 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.app.zul.DisposedEventListener;

/**
 * @author Ian Tsai
 *
 */
public class AbstractBaseContext implements BaseContext {
	/**
	 * @author Ian Tsai
	 */
	protected static class EventListenerStore{
		private Map<String, List<EventListener>> map = 
			new HashMap<String, List<EventListener>>();
		/**
		 * @param eventName
		 * @param eventListener
		 */
		public void add(String eventName, EventListener eventListener) {
			 List<EventListener> listeners = getEventListeners(eventName);
			 listeners.add(eventListener);
		}

		/**
		 * @param eventName
		 * @param eventListener
		 */
		public void remove(String eventName, EventListener eventListener) {
			 List<EventListener> listeners = getEventListeners(eventName);
			 if(listeners.size()==0)return;
			 listeners.remove(eventListener);
		}

		/**
		 * @param string
		 * @return
		 */
		public synchronized List<EventListener> getEventListeners(String eventName) {
			List<EventListener> listener = map.get(eventName);
			if(listener==null){
				map.put(eventName, listener = new LinkedList<EventListener>());
			}
			return listener;
		}
		
		public void fire( Event event){
			List<EventListener> listeners = 
				this.getEventListeners(event.getName());
			
			for(EventListener listener : new ArrayList<EventListener>(listeners)){
				
				if(listener instanceof DisposedEventListener){
					if(((DisposedEventListener)listener).isDisposed()){
						this.remove(event.getName(), listener);
						break;
					}
				}
				try {
					listener.onEvent(event);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		
	}//end of class...
	
	protected EventListenerStore listenerStore = new EventListenerStore();

	public void addEventListener(String eventName, EventListener eventListener) {
		listenerStore.add(eventName, eventListener);
	}
	
	public void removeEventListener(String eventName, EventListener eventListener) {
		listenerStore.remove(eventName, eventListener);
	}

}
