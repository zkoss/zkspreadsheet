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

/**
 * @author henri
 * @since 3.9.0
 */
public class DatesMatch implements Matchable<Date>, Serializable {
	private static final long serialVersionUID = 3180841702616760127L;
	final private int min;
	final private int max;
	public DatesMatch(int min, int max) {
		this.min = min;
		this.max = max;
	}
	@Override
	public boolean match(Date value) {
		final double d = DateUtil.getExcelDate(value);
		return min <= d && d < max;
	}
}
