/* SSDataListener.java

	Purpose:
		
	Description:
		
	History:
		Mar 30, 2010 4:28:02 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine;

import org.zkoss.zk.ui.event.EventListener;

/**
 * ZK Spreadsheet model event dispatcher(used for model to notify UI).
 * @author henrichen
 *
 */
public interface EventDispatcher {
	/**
	 * Add a spreadsheet data event listener to this event dispatcher .
	 * @param name event name
	 * @param listener event listener.
	 * @return true if succeed.
	 */
	public boolean addEventListener(String name, EventListener listener);

	/**
	 * Remove a spreadsheet data event listener from this event dispatcher.
	 * @param name name event name
	 * @param listener event listener
	 * @return true if succeed.
	 */
	public boolean removeEventListener(String name, EventListener listener);
}
