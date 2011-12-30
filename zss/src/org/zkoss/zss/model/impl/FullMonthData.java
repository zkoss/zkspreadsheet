/* FullWeekData.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 5:21:07 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

import java.text.DateFormatSymbols;
import java.util.Locale;

import org.zkoss.util.CacheMap;
import org.zkoss.util.Pair;

/**
 * @author henrichen
 *
 */
public class FullMonthData extends CircularData { //ZSS-69
	private FullMonthData(String[] data, int type, Locale locale) {
		super(data, type, locale);
	}
	private static final CacheMap _monthData;
	static {
		_monthData = new CacheMap(4);
		_monthData.setLifetime(24*60*60*1000);
	}
	public static FullMonthData getInstance(int type, Locale locale) {
		final Pair key = new Pair(locale, Integer.valueOf(type));
		FullMonthData value = (FullMonthData) _monthData.get(key);
		if (value == null) { //create and cache
			DateFormatSymbols symbols = DateFormatSymbols.getInstance(locale);
			if (symbols == null) {
				symbols = DateFormatSymbols.getInstance(Locale.US);
			}
			value = new FullMonthData(symbols.getMonths(), type, locale);
			_monthData.put(key, value);
		}
		return value;
	}
}
