/* BookSeries.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 12:04:14 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

/**
 * Represent a series of correlated {@link Book}s that might refer to each other.
 * <p>Note: this feature requires ZK Spreadsheet EE.</p>
 * 
 * @author henrichen
 * @author dennischen
 * @deprecated since 3.0.0, please use class in package {@code org.zkoss.zss.api}
 */
public interface BookSeries {
	/**
	 * Return the {@link Book} with the specified book name.
	 * @param bookName the book name
	 * @return the {@link Book} with the specified book name.
	 */
	public Book getBook(String bookName);
}
