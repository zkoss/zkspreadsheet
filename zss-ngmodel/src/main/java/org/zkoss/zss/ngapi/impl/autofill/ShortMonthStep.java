/* FullWeekStep.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 2:37:53 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.ngapi.impl.autofill;

import java.util.Locale;

import org.zkoss.poi.ss.usermodel.Cell;

/**
 * Full week Step.
 * @author henrichen
 *
 */
/*package*/ class ShortMonthStep implements Step {
	private final CircularStep _innerStep;
	private final int _type;
	public ShortMonthStep(int initial, int step, int type, int datatype, Locale locale) { //ZSS-69
		_innerStep = new CircularStep(initial, step, ShortMonthData.getInstance(type, locale));
		_type = datatype;
	}
	@Override
	public int getDataType() {
		return _type;
	}
	@Override
	public Object next(Cell srcCell) {
		return _innerStep.next(srcCell);
	}
}
