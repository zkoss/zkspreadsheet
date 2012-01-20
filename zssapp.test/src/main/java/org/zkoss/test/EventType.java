/* Event.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 5:49:53 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

/**
 * @author sam
 *
 */
public enum EventType {
	KEYDOWN, KEYUP, MOUSEDOWN, MOUSEUP, MOUSEOVER, MOUSEMOVE, MOUSECLICK, CONTEXTMENU;

	@Override
	public String toString() {
		return super.toString().toLowerCase();
	}
}
