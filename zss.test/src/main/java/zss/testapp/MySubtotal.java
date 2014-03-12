package zss.testapp;

import org.zkoss.poi.ss.formula.eval.EvaluationException;
import org.zkoss.poi.ss.formula.functions.MultiOperandNumericFunction;

/**
 * Implement a numeric function that accepts multiple operations
 * @author Hawk
 *
 */
public class MySubtotal extends MultiOperandNumericFunction{

	protected MySubtotal() {
		// the first parameter determines whether to evaluate boolean value. If
		// it's true, evaluator will evaluate boolean value to number. TRUE to 1
		// and FALSE to 0.
		// If it's false, boolean value is just ignored.
		// The second parameter determines whether to evaluate blank value.
		super(false, false);
	}
	
	/**
	 * inherited method, ValueEval evaluate(ValueEval[] args, int srcCellRow,
	 * int srcCellCol), will call this overridden method This function depends
	 * on MultiOperandNumericFunction that evaluates all arguments to double.
	 * This function sums all double values in cells.
	 */
	@Override
	protected double evaluate(double[] values) throws EvaluationException {
		double sum = 0;
		for (int i = 0 ; i < values.length ; i++){
			sum += values[i];
		}
		return sum;
	}
}
