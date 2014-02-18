/* LinearStep.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 2:29:38 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.ngapi.impl.autofill;

import org.zkoss.poi.ss.usermodel.Cell;

/**
 * Linear incremental Step.
 * @author henrichen
 * @since 2.1.0
 */
public class LinearStep implements Step {
	private double _current;
	private final double _step;
	private final int _type;
	/*package*/ LinearStep(double initial, double initStep, double step, int type) {
		_current = initial + initStep;
		_step = step;
		_type = type;
	}
	
	@Override
	public int getDataType() {
		return _type;
	}

	@Override
	public Object next(Cell cell) {
		if (cell.getCellType() != Cell.CELL_TYPE_NUMERIC) {
			return null;
		}
		final double current = _current;
		_current += _step;
		return current;
	}
}
