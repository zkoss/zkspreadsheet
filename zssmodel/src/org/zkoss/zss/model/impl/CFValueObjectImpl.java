/* CfValueObjectImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 12:50:53 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCFValueObject;

/**
 * @author henri
 * @since 3.8.2
 */
public class CFValueObjectImpl implements SCFValueObject, Serializable {
	private static final long serialVersionUID = 189125364763803062L;
	
	private CFValueObjectType type;
	private String value;
	private boolean gte;
	
	@Override
	public CFValueObjectType getType() {
		return type;
	}
	
	public void setType(CFValueObjectType type) {
		this.type = type;
	}

	@Override
	public String getValue() {
		return value;
	}
	
	public void setValue(String value) {
		this.value = value; 
	}

	@Override
	public boolean isGreaterOrEqual() {
		return gte;
	}
	
	public void setGreaterOrEqual(boolean b) {
		gte = b;
	}
}
