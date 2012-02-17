/* SheetCreateEvent.java

	Purpose:
		
	Description:
		
	History:
		Feb 7, 2012 11:51:32 AM, Created by henri

Copyright (C) 2012 Potix Corporation. All Rights Reserved.
*/
package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;

/**
 * When a sheet is created. 
 * @author henri
 */
public class SheetCreateEvent extends Event {
	private String _sheetName;
	public SheetCreateEvent(String name, Component target, String sheetName) {
		super(name, target);
		_sheetName = sheetName;
	}
	public String getSheetName() {
		return _sheetName;
	}
}
