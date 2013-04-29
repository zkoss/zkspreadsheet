/* ModelException.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Wed Jun 13 16:41:56     2007, Created by tomyeh
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.model.sys;

import org.zkoss.lang.SystemException;

/**
 * Represents a data model related exception.
 *
 * @author tomyeh
 */
public class ModelException extends SystemException {
	private static final long serialVersionUID = 8278035139743537846L;

	public ModelException(String msg, Throwable cause) {
		super(msg, cause);
	}
	public ModelException(String s) {
		super(s);
	}
	public ModelException(Throwable cause) {
		super(cause);
	}
	public ModelException() {
	}

	public ModelException(int code, Object[] fmtArgs, Throwable cause) {
		super(code, fmtArgs, cause);
	}
	public ModelException(int code, Object fmtArg, Throwable cause) {
		super(code, fmtArg, cause);
	}
	public ModelException(int code, Object[] fmtArgs) {
		super(code, fmtArgs);
	}
	public ModelException(int code, Object fmtArg) {
		super(code, fmtArg);
	}
	public ModelException(int code, Throwable cause) {
		super(code, cause);
	}
	public ModelException(int code) {
		super(code);
	}
}
