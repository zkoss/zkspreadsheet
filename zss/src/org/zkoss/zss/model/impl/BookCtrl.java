/* BookCtrl.java

	Purpose:
		
	Description:
		
	History:
		Dec 7, 2010 11:18:37 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.model.Book;

/**
 * Book controls (Internal Use only).
 * @author henrichen
 *
 */
public interface BookCtrl {
	/**
	 * Create an associated reference book.
	 * @param book the ZK Spreadsheet book
	 */
	public RefBook newRefBook(Book book);
}
