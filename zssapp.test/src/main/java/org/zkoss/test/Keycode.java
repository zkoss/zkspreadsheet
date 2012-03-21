/* Keyboard.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 5:58:40 PM , Created by sam
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
public enum Keycode {
	ENTER(13), 
	ESC(27),
	DELETE(46),
	A(65),
	B(66),
	C(67),
	D(68),
	E(69),
	I(73),
	O(79),
	U(85),
	V(86),
	X(88);
	
	private final int keyCode;
	Keycode(int keyCode) {
		this.keyCode = keyCode;
	}
	
	public int intValue() {
		return keyCode;
	}
	
	//TODO: public List<Keycode> valueOf(String input) {}
}
