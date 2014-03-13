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
package org.zkoss.zss.model;
/**
 * This exception is thrown when you perform an unreasonable operation on the model.
 * @author dennis
 * @since 3.5.0
 */
public class InvalidateModelOpException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public InvalidateModelOpException() {
		super();
	}

	public InvalidateModelOpException(String message, Throwable cause) {
		super(message, cause);
	}

	public InvalidateModelOpException(String message) {
		super(message);
	}

	public InvalidateModelOpException(Throwable cause) {
		super(cause);
	}

}
