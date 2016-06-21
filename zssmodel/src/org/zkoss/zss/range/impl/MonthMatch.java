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
public class MonthMatch implements Matchable<Date>, Serializable {
	final private int base;
	public MonthMatch(int base) {
		this.base = base;
	}
	@Override
	public boolean match(Date value) {
		final Calendar now = new GregorianCalendar();
		now.setTimeInMillis(value.getTime());
		final int v = now.get(Calendar.MONTH);
		return v == base;
	}
}
