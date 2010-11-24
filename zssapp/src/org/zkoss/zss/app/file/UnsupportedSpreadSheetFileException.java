/* UnsupportedSpreadSheetFileException.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 3, 2010 3:04:22 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

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
