/* IllegalOpArgumentException.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/7/24 , Created by kuro
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api;

/**
 * Indicate a illegal operation & argument exception, 
 * usually, we catch this exception and show corresponding message for user 
 * @author kuro
 * @since 3.0.0
 */
public class IllegalOpArgumentException extends RuntimeException {
	private static final long serialVersionUID = 1L;
	
	public IllegalOpArgumentException() {
		super();
	}

	public IllegalOpArgumentException(String message, Throwable cause) {
		super(message, cause);
	}

	public IllegalOpArgumentException(String message) {
		super(message);
	}

	public IllegalOpArgumentException(Throwable cause) {
		super(cause);
	}
}
