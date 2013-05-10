/* DateStep.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 1, 2011 9:37:35 AM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.sys.impl;

import java.util.Calendar;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.DateUtil;

/**
 * Date step.
 * @author henrichen
 * @since 2.1.0
 */
public class DateStep implements Step {
	private final Calendar _cal = Calendar.getInstance(); //TODO timezone?
	private final int _year;
	private final int _month;
	private final int _date;
	private final int _time;
	
	private final int _mStep;
	private final int _tStep;
	
	private int _steps;
	private final int _type;
	private final long _zero;
	
	public DateStep(int y, int m, int d, int t, int mStep, int tStep, int type) {
		_year = y;
		_month = m;
		_date = d;
		_time = t;
		
		_mStep = mStep;
		_tStep = tStep;
		
		_steps = 0;
		_type = type;
		_zero = DateUtil.getJavaDate(0.0).getTime();
	}
	
	public Object next(Cell cell) {
		++_steps;
		
		_cal.clear();
		_cal.set(_year, _month + _mStep * _steps, 1);
		final int date = Math.min(_date, _cal.getActualMaximum(Calendar.DAY_OF_MONTH)); 
		_cal.set(Calendar.DAY_OF_MONTH, date);
		
		long ms = _cal.getTime().getTime() + _time + _tStep * _steps;
		if (ms < _zero) {
			ms += 24 * 60 * 60 * 1000L;
		}
		_cal.setTimeInMillis(ms);
		
		return _cal.getTime();
	}

	@Override
	public int getDataType() {
		return _type;
	}
}
