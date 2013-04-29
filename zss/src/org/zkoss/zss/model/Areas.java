/* Areas.java

	Purpose:
		
	Description:
		
	History:
		Oct 28, 2010 2:52:51 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;


/**
 * A collection of single-area {@link Range}s in a selection; an area is a contiguous block of cells.
 * @author henrichen
 * @author dennischen
 * @deprecated since 3.0.0, please use class in package {@code org.zkoss.zss.api}
 */
public interface Areas extends Iterable<Range> {
	/**
	 * Returns the number of single area Range in this collection.
	 * @return the number of objects in this collection.
	 */
	public int getCount();
	
	/**
	 * Returns a single object from this collection. 
	 */
	public Range getItem();
}

