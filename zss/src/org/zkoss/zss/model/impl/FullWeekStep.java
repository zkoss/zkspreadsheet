/* FullWeekStep.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 2:37:53 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

import org.zkoss.poi.ss.usermodel.Cell;

/**
 * Full week Step.
 * @author henrichen
 *
 */
/*package*/ class FullWeekStep implements Step {
	private CircularStep _innerStep;
	public FullWeekStep(int initial, int step, int type) {
		_innerStep = new CircularStep(initial, step, new FullWeekData(type));
	}
	@Override
	public Object next(Cell srcCell) {
		return _innerStep.next(srcCell);
	}
}
