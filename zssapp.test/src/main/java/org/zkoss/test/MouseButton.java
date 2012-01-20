/* Mouse.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 4:22:26 PM , Created by sam
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
public enum MouseButton {
	IGNORE(-1), LEFT(1), MIDDLE(2), RIGHT(3);
	
	private int which;
	private MouseButton(int which) {
		this.which = which;
	}
	
	public int getWhich() {
		return which;
	}
}
