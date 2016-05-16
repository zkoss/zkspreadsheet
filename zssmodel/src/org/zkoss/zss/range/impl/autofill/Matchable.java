/* Matchable.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 7:37:42 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl.autofill;

/**
 * @author henri
 * @since 3.9.0
 */
public interface Matchable<T extends Comparable<T>> {
	boolean match(T value);
}
