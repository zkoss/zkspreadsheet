/* NBookSeries.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/11/14 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.util.concurrent.locks.ReadWriteLock;

/**
 * @author dennis
 *
 */
public interface NBookSeries {
	
	/**
	 * Get the book by name;
	 * @param name the book name
	 * @return the book or null if not found.
	 */
	public NBook getBook(String name);
	
	/**
	 * Get the ReadWriteLock for synchronized when read-write model for current accessing.
	 * @return
	 */
	public ReadWriteLock getLock();
}
