/* CellStyleContextEvent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 10:31:40 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.zk.ui.event.Event;

/**
 * @author Sam
 *
 */
public class CellStyleContextEvent extends Event {

	
	private Object executor; 
	/**
	 * @param name
	 * @param target
	 * @param data
	 */
	public CellStyleContextEvent(String eventName, CellStyle data) {
		super(eventName, null, data);
	}
	

	public CellStyle getCellStyle() {
		return (CellStyle) getData();
	}


	public Object getExecutor() {
		return executor;
	}

	public void setExecutor(Object executor) {
		this.executor = executor;
	}
	
}