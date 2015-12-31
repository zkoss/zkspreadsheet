/* IconSetImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 2:23:13 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SIconSet;

/**
 * @author henri
 * @since 3.8.2
 */
public class IconSetImpl implements SIconSet {
	private IconSetType type;
	private List<SCFValueObject> valueObjects;
	private boolean percent;
	private boolean reverse;
	private boolean showValue = true; //ZSS-1161: default true
	
	@Override
	public IconSetType getType() {
		return type;
	}
	
	public void setType(IconSetType type) {
		this.type = type;
	}

	@Override
	public List<SCFValueObject> getCFValueObjects() {
		return valueObjects;
	}
	
	public void addValueObject(SCFValueObject vo) {
		if (valueObjects == null) {
			valueObjects = new ArrayList<SCFValueObject>();
		}
		valueObjects.add(vo);
	}

	@Override
	public boolean isPercent() {
		return percent;
	}
	
	public void setPercent(boolean b) {
		percent = b;
	}

	@Override
	public boolean isReverse() {
		return reverse;
	}

	public void setReverse(boolean b) {
		reverse = b;
	}
	
	@Override
	public boolean isShowValue() {
		return showValue;
	}
	
	public void setShowValue(boolean b) {
		showValue = b;
	}
}
