/* DataEventDispatchListener.java

	Purpose:
		
	Description:
		
	History:
		Mar 12, 2010 1:08:20 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.event;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.engine.EventDispatcher;

/**
 * Generic spreadsheet data event dispatch listener(used for dispatching
 * spreadsheet data events from the spreadsheet Book to spreadsheet UI).
 * Each spreadsheet UI must has exactly one associated DataEventDispatchListener.
 * 
 * @author henrichen
 */
public class EventDispatchListener implements EventListener, EventDispatcher {
	private Map<String, List<EventListener>> _listeners = new HashMap<String, List<EventListener>>(8);

	//--EventListener--//
	@Override
	public void onEvent(Event event) throws Exception {
		final String name = event.getName();
		final List<EventListener> list = _listeners.get(name);
		if (list != null) {
			for(EventListener listener : list) {
				listener.onEvent(event);
			}
		}
	}

	//--EventDispatcher--//
	@Override
	public boolean addEventListener(String name, EventListener listener) {
		List<EventListener> list = _listeners.get(name);
		if (list == null) {
			list = new ArrayList<EventListener>(4);
			_listeners.put(name, list);
		}
		return list.add(listener);
	}

	@Override
	public boolean removeEventListener(String name, EventListener listener) {
		final List<EventListener> list = _listeners.get(name);
		return list != null ? list.remove(listener) : null;
	}
}
