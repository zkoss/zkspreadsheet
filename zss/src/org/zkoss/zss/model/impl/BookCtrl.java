/* BookCtrl.java

	Purpose:
		
	Description:
		
	History:
		Dec 7, 2010 11:18:37 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.List;

import org.zkoss.poi.ss.usermodel.PivotCache;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.model.Book;

/**
 * Book controls (Internal Use only).
 * @author henrichen
 *
 */
public interface BookCtrl {
	public static final String CLASS = "org.zkoss.zss.model.BookCtrl.class";
	/**
	 * Create an associated reference book.
	 * @param book the ZK Spreadsheet book
	 */
	public RefBook newRefBook(Book book);
	
	/**
	 * Return next sheet id.
	 * @return next sheet id.
	 */
	public Object nextSheetId();
	
	/**
	 * Return next focus id for the UI.
	 * @return next focus id.
	 */
	public String nextFocusId();
	
	public void addFocus(Object focus);
	
	public void removeFocus(Object focus);
	
	public boolean containsFocus(Object focus);
}
