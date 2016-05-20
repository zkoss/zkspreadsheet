/* FilterRowInfo.java

	Purpose:
		
	Description:
		
	History:
		May 19, 2016 5:01:30 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;

import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
import org.zkoss.zss.model.ErrorValue;

/**
 * @author henri
 * @since 3.9.0
 */
public class FilterRowInfo implements Serializable {
	private static final long serialVersionUID = 8996620315976386725L;
	
	private Object value;
	private String display;
	private boolean seld;
	
	public FilterRowInfo(Object val, String displayVal) {
		value = val;
		display = displayVal;
	}
	
	public Object getValue() {
		return value;
	}
	
	public String getDisplay() {
		return display;
	}
	
	public void setSelected(boolean selected) {
		seld = selected;
	}
	
	public boolean isSelected() {
		return seld;
	}

	public int hashCode() {
		return value == null ? 0 : value.hashCode();
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof FilterRowInfo))
			return false;
		final FilterRowInfo other = (FilterRowInfo) obj;
		return Objects.equals(other.value, this.value);
	}

	//ZSS-1191
	//Date < Number < String < Boolean(FALSE < TRUE) < Error(byte) < (Blanks)
	public  static int getType(Object val) {
		if (val instanceof Date) {
			return 1;
		}
		if (val instanceof ErrorValue || val instanceof Byte) { //error, ZSS-707
			return 5;
		}
		if (val instanceof Number) {
			return 2;
		}
		if (val instanceof String) {
			return Strings.isEmpty((String)val) ? 6 : 3;
		}
		if (val instanceof Boolean) {
			return 4;
		}
		return 6;
	}
}