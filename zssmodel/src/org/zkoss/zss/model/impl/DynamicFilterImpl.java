/* DynamicFilterImpl.java

	Purpose:
		
	Description:
		
	History:
		May 16, 2016 6:34:01 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SDynamicFilter;

/**
 * DynamicFilter
 * @author henri
 *
 */
public class DynamicFilterImpl implements SDynamicFilter, Serializable {
	public static final SDynamicFilter NOOP_DYNAFILTER = new DynamicFilterImpl(null, null, true); //ZSS-1193

	private Double maxValue;
	private Double value;
	private boolean isAbove;

	public DynamicFilterImpl(Double maxValue, Double value, boolean isAbove) {
		this.maxValue = maxValue;
		this.value = value;
		this.isAbove = isAbove;
	}
	
	@Override
	public Double getMaxValue() {
		return maxValue;
	}
	
	@Override
	public Double getValue() {
		return value;
	}
	
	@Override
	public boolean isAbove() {
		return isAbove;
	}
	
	DynamicFilterImpl cloneDynamicFilter() {
		return new DynamicFilterImpl(this.maxValue, this.value, this.isAbove);
	}
}
