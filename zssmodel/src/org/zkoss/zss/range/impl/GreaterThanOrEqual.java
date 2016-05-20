/* GreaterThanOrEqual.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:15:30 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;

/**
 * @author henri
 * @since 3.9.0
 */
public class GreaterThanOrEqual<T extends Comparable<T>> implements Matchable<T>, Serializable {
	private static final long serialVersionUID = 7460476308225284031L;
	
	final private T base;
	public GreaterThanOrEqual(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) >= 0;
	}
}

