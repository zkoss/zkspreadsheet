package org.zkoss.zss.ui;

import java.util.List;
import java.util.Set;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.SerializableEventListener;
import org.zkoss.zss.api.model.Sheet;
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
	 * @return event name list if you have any interested event of spreadsheet.
	 * @see Events
	 */
	Set<String> getInterestedEvents();
	
	
	/**
	 * Returns the interested ctrlKeys of the spreadsheet
	 * @return ctrlKeys that you want to set to spreadsheet, or null to set nothing to spreadsheet.
	 * @see Spreadsheet#setCtrlKeys(String)
	 */
	String getCtrlKeys();
	
	
	/**
	 * Returns the supported user action that should be disabled 
	 * @param sheet the sheet for cheeking
	 * @return a disabled user action array
	 */
	Set<String> getSupportedUserAction(Sheet sheet);
	
	
	/**
	 * Sets the spreadsheet this handler relates to. this method is called when it assign to a spreadsheet 
	 * @param sparedsheet
	 */
	void bind(Spreadsheet sparedsheet);
	
//	/**
//	 * the i18n label keys for client side
//	 * @return
//	 */
//	String[] getLabelKeys();
}
