/* LessThanOrEqual.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:19:03 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;

/**
 * @author henri
 * @since 3.9.0
 */
public class LessThanOrEqual<T extends Comparable<T>> implements Matchable<T>, Serializable {
	private static final long serialVersionUID = 1L;
	
	final private T base;
	public LessThanOrEqual(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) <= 0;
	}
}
