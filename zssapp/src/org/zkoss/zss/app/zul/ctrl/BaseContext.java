/* BaseContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 6:19:33 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.zk.ui.event.EventListener;

/**
 * @author Sam
 *
 */
public interface BaseContext {
	/**
	 * 
	 * @param evtName
	 * @param eventListener
	 * @return
	 */
	public void addEventListener(String evtName, EventListener eventListener);
	/**
	 * 
	 * @param evtName
	 * @param eventListener
	 * @return
	 */
	public void removeEventListener(String evtName, EventListener eventListener);
}
