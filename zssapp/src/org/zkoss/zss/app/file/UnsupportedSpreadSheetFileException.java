/* UnsupportedSpreadSheetFileException.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 3, 2010 3:04:22 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.file;

/**
 * @author Sam
 *
 */
public class UnsupportedSpreadSheetFileException extends Exception {

	/**
	 * 
	 */
	public UnsupportedSpreadSheetFileException() {
	}

	/**
	 * @param arg0
	 */
	public UnsupportedSpreadSheetFileException(String arg0) {
		super(arg0);
	}

	/**
	 * @param arg0
	 */
	public UnsupportedSpreadSheetFileException(Throwable arg0) {
		super(arg0);
	}

	/**
	 * @param arg0
	 * @param arg1
	 */
	public UnsupportedSpreadSheetFileException(String arg0, Throwable arg1) {
		super(arg0, arg1);
	}

}
