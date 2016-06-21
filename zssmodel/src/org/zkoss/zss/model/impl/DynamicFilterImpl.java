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
	public static final SDynamicFilter NOOP_DYNAFILTER = new DynamicFilterImpl(null, null, null); //ZSS-1193, ZSS-1234

	private String type; //ZSS-1234
	private Double maxValue;
	private Double value;

	public DynamicFilterImpl(Double maxValue, Double value, String type) {
		this.maxValue = maxValue;
		this.value = value;
		this.type = type == null ? "null" : type;
	}
	
	//ZSS-1234
	@Override
	public String getType() {
		return type;
	}
	
	@Override
	public Double getMaxValue() {
		return maxValue;
	}
	
	@Override
	public Double getValue() {
		return value;
	}
	
	DynamicFilterImpl cloneDynamicFilter() {
		return new DynamicFilterImpl(this.maxValue, this.value, this.type); //ZSS-1234
	}
}
