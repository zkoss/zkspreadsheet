/* Top10FilterImpl.java

	Purpose:
		
	Description:
		
	History:
		May 17, 2016 4:57:03 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.STop10Filter;

//ZSS-1227
/**
 * @author henri
 * @since 3.9.0
 */
public class Top10FilterImpl implements STop10Filter, Serializable {
	private static final long serialVersionUID = 1417793891432418239L;
	
	private boolean isTop;
	private boolean isPercent;
	private double value;
	private double filterValue;
	
	public Top10FilterImpl(boolean isTop, double value, boolean isPercent, double filterValue) {
		this.isTop = isTop;
		this.isPercent = isPercent;
		this.value = value;
		this.filterValue = filterValue;
	}
	
	@Override
	public boolean isTop() {
		return isTop;
	}

	@Override
	public boolean isPercent() {
		return isPercent;
	}

	@Override
	public double getValue() {
		return value;
	}

	@Override
	public double getFilterValue() {
		return filterValue;
	}

	Top10FilterImpl cloneTop10Filter() {
		return new Top10FilterImpl(this.isTop, this.value, this.isPercent, this.filterValue);
	}
}
