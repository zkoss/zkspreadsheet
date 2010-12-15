/* BookCtrlImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 7, 2010 11:30:44 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.impl.RefBookImpl;
import org.zkoss.zss.model.Book;

/**
 * Implementation of {@link BookCtrl}.
 * @author henrichen
 *
 */
public class BookCtrlImpl implements BookCtrl {
	private int _shid;
	
	@Override
	public RefBook newRefBook(Book book) {
		return new RefBookImpl(book.getBookName(), book.getSpreadsheetVersion().getLastRowIndex(), book.getSpreadsheetVersion().getLastColumnIndex());
	}
	
	public Object nextSheetId() {
		return Integer.toString((_shid++ & 0x7FFFFFFF), 32);
	}
}
