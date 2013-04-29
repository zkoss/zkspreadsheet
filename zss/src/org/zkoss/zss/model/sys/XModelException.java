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
public class XModelException extends SystemException {
	private static final long serialVersionUID = 8278035139743537846L;

	public XModelException(String msg, Throwable cause) {
		super(msg, cause);
	}
	public XModelException(String s) {
		super(s);
	}
	public XModelException(Throwable cause) {
		super(cause);
	}
	public XModelException() {
	}

	public XModelException(int code, Object[] fmtArgs, Throwable cause) {
		super(code, fmtArgs, cause);
	}
	public XModelException(int code, Object fmtArg, Throwable cause) {
		super(code, fmtArg, cause);
	}
	public XModelException(int code, Object[] fmtArgs) {
		super(code, fmtArgs);
	}
	public XModelException(int code, Object fmtArg) {
		super(code, fmtArg);
	}
	public XModelException(int code, Throwable cause) {
		super(code, cause);
	}
	public XModelException(int code) {
		super(code);
	}
}
