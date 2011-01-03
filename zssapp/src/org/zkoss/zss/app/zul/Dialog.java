/* Dialog.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 6:22:52 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.event.Events;
import org.zkoss.zul.Window;

/**
 * Dialog doesn't detach when close
 * @author sam
 *
 */
public class Dialog extends Window {
	
	public void onCancel() {
		setVisible(false);
	}
	
	public void onClose() {
		setVisible(false);
	}
	
	/**
	 * open dialog and fire onOpen event
	 * @param obj data of event
	 */
	public void fireOnOpen(Object obj) {
		setVisible(true);
		Events.sendEvent("onOpen", this, obj);
	}
	
	public void close() {
		setVisible(false);
	}
}
