/* CustomFilterImpl.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 11:51:42 AM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCustomFilter;

//ZSS-1224
/**
 * @author henri
 * @since 3.9.0
 */
public class CustomFilterImpl implements SCustomFilter, Serializable {
	private String value;
	private SCustomFilter.Operator operator;
	
	public CustomFilterImpl(String value, Operator operator) {
		this.value = value;
		this.operator = operator;
	}
	
	@Override
	public String getValue() {
		return value;
	}

	@Override
	public Operator getOperator() {
		// TODO Auto-generated method stub
		return operator;
	}

	//ZSS-1183, ZSS-1224
	/*package*/ CustomFilterImpl cloneCustomFilter() {
		return new CustomFilterImpl(value, operator);
	}
}
