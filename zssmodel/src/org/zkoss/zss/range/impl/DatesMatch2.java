/* GreaterThanOrEqual.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:15:30 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Date;

import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.zss.model.SCell;

/**
 * @author henri
 * @since 3.9.0
 */
public class DatesMatch2 implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = 3180841702616760127L;
	final private int min;
	final private int max;
	public DatesMatch2(int min, int max) {
		this.min = min;
		this.max = max;
	}
	@Override
	public boolean match(SCell cell) {
		final Object value = cell.getValue();
		if (value instanceof Date || value instanceof Number) {
			final double d = 
					value instanceof Date ?
							DateUtil.getExcelDate((Date)value) : 
							((Number)value).doubleValue();
			return min <= d && d < max;
		} else {
			return false;
		}
	}
}
