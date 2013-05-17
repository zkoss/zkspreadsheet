package org.zkoss.zss.ui;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.SerializableEventListener;
import org.zkoss.zss.ui.event.Events;

/**
 * Action Handler for user's action
 * 
 * @author dennis
 * 
 */
public interface UserActionHandler extends SerializableEventListener<Event>{

	/**
	 * Returns the interested events of the spreadsheet.  
	 * @return event name array if you have any interested event of spreadsheet.
	 * @see Events
	 */
	String[] getInterestedEvents();
	
	
	/**
	 * Returns the interested ctrlKeys of the spreadsheet
	 * @return ctrlKeys that you want to set to spreadsheet, or null to set nothing to spreadsheet.
	 * @see Spreadsheet#setCtrlKeys(String)
	 */
	String getCtrlKeys();
	
//	/**
//	 * the i18n label keys for client side
//	 * @return
//	 */
//	String[] getLabelKeys();
}
