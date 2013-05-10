/* BookSeries.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 12:04:14 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys;

/**
 * Represent a series of correlated {@link XBook}s that might refer to each other.
 * <p>Note: this feature requires ZK Spreadsheet EE.</p>
 * 
 * @author henrichen
 */
public interface XBookSeries {
	/**
	 * Return the {@link XBook} with the specified book name.
	 * @param bookName the book name
	 * @return the {@link XBook} with the specified book name.
	 */
	public XBook getBook(String bookName);
}
