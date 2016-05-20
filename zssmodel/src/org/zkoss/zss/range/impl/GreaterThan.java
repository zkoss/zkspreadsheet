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
 * @author henri
 * @since 3.9.0
 */
public class GreaterThan<T extends Comparable<T>> implements Matchable<T>, Serializable {
	private static final long serialVersionUID = 378067185714337866L;
	
	final private T base;
	public GreaterThan(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) > 0;
	}
}
