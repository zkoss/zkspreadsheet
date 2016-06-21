/* GreaterThan.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:13:53 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;

/**
 * 
 * @author henri
 * @since 3.9.0
 */
public class GreaterThanOrEqualAndLessThan<T extends Comparable<T>> implements Matchable<T>, Serializable {
	
	final private T min;
	final private T max;
	// between min and max where min inclusive; max exclusive
	public GreaterThanOrEqualAndLessThan(T min, T max) {
		this.min = min;
		this.max = max;
	}
	
	@Override
	public boolean match(T value) {
		return value.compareTo(min) >= 0 && value.compareTo(max) < 0;
	}
}
