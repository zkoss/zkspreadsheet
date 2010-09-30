/* SheetSelector.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 27, 2010 7:44:25 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.zkoss.zss.model.Book;

/**
 * @author Sam
 *
 */
public class SheetSelector {
	
	/**
	 * Visit each sheet in book
	 */
	public void doVisit(Book book, SheetVisitor visitor) {
		int numSheet = book.getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			visitor.handle(book.getSheetAt(i));
		}
	}
}
