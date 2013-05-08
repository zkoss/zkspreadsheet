package org.zkoss.zss.ui;

import java.util.EventListener;
import java.util.Map;

import org.zkoss.zss.api.model.Sheet;

/**
 * Action Handler for user's {@code Action}
 * 
 * @author dennis
 * 
 */
public interface UserActionHandler {

	/**
	 * handle user's action
	 * @param spreadsheet the spreadsheet
	 * @param targetSheet the target sheet for the action
	 * @param action the action
	 * @param selection the selection 
	 * @param extraData the extra data for the action
	 * @return true if the action is handled, false if not.
	 */
	boolean handleAction(Spreadsheet spreadsheet, Sheet sheet,String action,
			Rect selection, Map<String, Object> extraData);
	
	/**
	 * Return the interested events of the spreadsheet. 
	 * Note: you have to also implement {@link EventListener} as the callback if you doesn't return null 
	 * @return event name array if you have any interested event of spreadsheet.
	 */
	String[] getInterestedEvents();
}
