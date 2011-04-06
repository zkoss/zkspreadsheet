/* MsecondStep.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 1, 2011 9:31:44 AM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

import java.util.Calendar;
import java.util.Date;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.DateUtil;

/**
 * Millisecond Step.
 * @author henrichen
 *
 */
public class MsecondStep implements Step {
	private final Calendar _cal = Calendar.getInstance(); //TODO timezone?
	private long _current;
	private final long _step;
	private final int _type;
	private final long _zero;
	
	public MsecondStep(Date date, long step, int type) {
		_current = date.getTime();
		_step = step;
		_type = type;
		_zero = DateUtil.getJavaDate(0.0).getTime();
	}
	@Override
	public Object next(Cell cell) {
		_current += _step;
		if (_current < _zero) {
			_current += 24 * 60 * 60 * 1000L;
		}
		_cal.setTimeInMillis(_current);
		return _cal.getTime();
	}
	@Override
	public int getDataType() {
		return _type;
	}
}
