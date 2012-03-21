/* CtrlKeyCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 21, 2012 10:25:09 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;

import java.util.Map;

import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.au.AuRequests;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.ui.event.KeyEvent;

/**
 * @author sam
 *
 */
public class CtrlKeyCommand implements Command {

	@Override
	public void process(AuRequest request) {
		Map data = request.getData();
		
		org.zkoss.zk.ui.event.KeyEvent evt = KeyEvent.getKeyEvent(request);
		Event zssKeyEvt = new KeyEvent(evt.getName(), evt.getTarget(), 
				evt.getKeyCode(), evt.isCtrlKey(), evt.isShiftKey(), evt.isAltKey(), 
				AuRequests.getInt(data, "tRow", -1), AuRequests.getInt(data, "lCol", -1),
				AuRequests.getInt(data, "bRow", -1), AuRequests.getInt(data, "rCol", -1));
		Events.postEvent(zssKeyEvt);
	}
}
