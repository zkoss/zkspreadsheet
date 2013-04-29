/* Areas.java

	Purpose:
		
	Description:
		
	History:
		Oct 28, 2010 2:52:51 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys;


/**
 * A collection of single-area {@link XRange}s in a selection; an area is a contiguous block of cells.
 * @author henrichen
 *
 */
public interface XAreas extends Iterable<XRange> {
	/**
	 * Returns the number of single area Range in this collection.
	 * @return the number of objects in this collection.
	 */
	public int getCount();
	
	/**
	 * Returns a single object from this collection. 
	 */
	public XRange getItem();
}

