/* Indexable.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 4:23:19 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

/**
 * Used in object with an index.
 * @author henrichen
 *
 */
public interface Indexable extends Comparable<Indexable> {
	/** Return the index */
	public int getIndex();
}
