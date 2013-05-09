package org.zkoss.zss.ui;

import java.util.EventListener;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.SerializableEventListener;

/**
 * Action Handler for user's action
 * 
 * @author dennis
 * 
 */
public interface UserActionHandler extends SerializableEventListener<Event>{

	/**
	 * Return the interested events of the spreadsheet. 
	 * Note: you have to also implement {@link EventListener} as the callback if you doesn't return null 
	 * @return event name array if you have any interested event of spreadsheet.
	 */
	String[] getInterestedEvents();
	
//	/**
//	 * the i18n label keys for client side
//	 * @return
//	 */
//	String[] getLabelKeys();
}
