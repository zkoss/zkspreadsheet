/* KeyEvent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 21, 2012 9:47:14 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zss.api.Rect;

/**
 * @author sam
 *
 */
public class KeyEvent extends org.zkoss.zk.ui.event.KeyEvent {

	final Rect selection;
	
	public KeyEvent(String name, Component target, int keyCode,
			boolean ctrlKey, boolean shiftKey, boolean altKey,
			int tRow, int lCol, int bRow, int rCol) {
		super(name, target, keyCode, ctrlKey, shiftKey, altKey);
		
		selection = new Rect(lCol, tRow, rCol, bRow);
	}
	
	public Rect getSelection() {
		return selection;
	}

}
