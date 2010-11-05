/* SheetVisitor.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 27, 2010 4:21:28 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.zkoss.poi.ss.usermodel.Sheet;

/**
 * @author Sam
 *
 */
public interface SheetVisitor {

	/**
	 * @param sheet
	 */
	public void handle(Sheet sheet);

}
