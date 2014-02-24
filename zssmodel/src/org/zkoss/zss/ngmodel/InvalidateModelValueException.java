/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class InvalidateModelValueException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public InvalidateModelValueException() {
		super();
	}

	public InvalidateModelValueException(String message, Throwable cause) {
		super(message, cause);
	}

	public InvalidateModelValueException(String message) {
		super(message);
	}

	public InvalidateModelValueException(Throwable cause) {
		super(cause);
	}

}
