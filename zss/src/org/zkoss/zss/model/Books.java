/* Books.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 12:04:14 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

/**
 * Represent multiple correlated books that might refer to each other.
 * @author henrichen
 */
public interface Books {
	/**
	 * Return the {@link Book} with the specified book name.
	 * @param bookName the book name
	 * @return the {@link Book} with the specified book name.
	 */
	public Book getBook(String bookName);
}
