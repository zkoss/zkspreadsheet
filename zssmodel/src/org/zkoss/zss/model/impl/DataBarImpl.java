/* DataBarImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 12:41:22 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SDataBar;

/**
 * @author henri
 *
 */
public class DataBarImpl implements SDataBar {
	private List<SCFValueObject> valueObjects;
	private SColor color;
	private Long minLength;
	private Long maxLength;
	private boolean showValue = true; //ZSS-1161: default to true
	
	@Override
	public List<SCFValueObject> getCFValueObjects() {
		return valueObjects;
	}
	
	public void addValueObject(SCFValueObject vobject) {
		if (valueObjects == null) {
			valueObjects = new ArrayList<SCFValueObject>();
		}
		valueObjects.add(vobject);
	}

	@Override
	public SColor getColor() {
		return color;
	}
	
	public void setColor(SColor color) {
		this.color = color;
	}

	@Override
	public Long getMinLength() {
		return minLength;
	}

	public void setMinLength(long minLength) {
		this.minLength = minLength;
	}
	
	@Override
	public Long getMaxLength() {
		return maxLength;
	}

	public void setMaxLength(long maxLength) {
		this.maxLength = maxLength;
	}
	
	@Override
	public boolean isShowValue() {
		return showValue;
	}
	
	public void setShowValue(boolean b) {
		showValue = b;
	}

}
