/* LinearStep.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 2:29:38 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

import org.zkoss.poi.ss.usermodel.Cell;

/**
 * Linear incremental Step.
 * @author henrichen
 * @since 2.1.0
 */
public class LinearStep implements Step {
	private double _current;
	private double _step;
	/*package*/ LinearStep(double initial, double initStep, double step) {
		_current = initial + initStep;
		_step = step;
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
	/*package*/ static Step getLinearStep(Cell[] srcCells, int b, int e, boolean positive) {
		int count = e-b+1;
		final double[] values = new double[count];
		for (int j = 0, k = b; k <=e; ++k) {
			final Cell srcCell = srcCells[k];
			values[j++] = srcCell.getNumericCellValue();
		}
		if (count == 1) {
			final double step = positive ? 1 : -1;
			return new LinearStep(values[count-1], step, step);
		} else if (count == 2) { //standard linear series
			final double step = values[1] - values[0];
			return new LinearStep(values[count-1], step, step);
		} else if (count == 3) { //3 source case (by experiment)
			double step = values[2] - values[0];
			double initStep	= (step + values[1] - values[0]) / 3;
			step /= 2;
			return new LinearStep(values[count-1], initStep, step);
		} else if (count == 4) { //4 source case (by experiment)
			double initStep = (values[2] - values[0]) / 2;
			double step = (values[3]-values[0]) * 0.3 + (values[2]-values[1]) * 0.1;
			return new LinearStep(values[count-1], initStep, step);
		}
		//TODO, for values equals to 5 or above 5, we apply the 5 values rule, though it is not the same to the Excel!
		//else if (j >= 5) { //5 source case (by experiment) 
			double initStep = -0.4 * values[0] - 0.1 * values[1] + 0.2 * values[2] + 0.5 * values[3] - 0.2 * values[4];
			double step = -0.2 * values[0] - 0.1 * values[1] + 0.1 * values[3] + 0.2 * values[4];
			return new LinearStep(values[count-1], initStep, step);
		//}
	}
}
