package org.zkoss.zss.essential.advanced;

import java.util.LinkedList;
import java.util.List;

import org.zkoss.poi.ss.formula.TwoDEval;
import org.zkoss.poi.ss.formula.eval.EvaluationException;
import org.zkoss.poi.ss.formula.eval.RefEval;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.functions.Function;
import org.zkoss.poi.ss.formula.functions.MultiOperandNumericFunction;

/**
 * Example of custom formulas. From basic to advanced one.
 * @author Hawk
 *
 */
public class MyCustomFormulas {

	/**
	 * Basic - Simple Formula.
	 * Exchange one money to another one according to specified exchange rate.
	 * @param money
	 * @param exchangeRate
	 * @return
	 */
	public static double exchange(double money, double exchangeRate) {
		return money * exchangeRate;
	}	

	/* 
	 * Intermediate - Multiple Numeric Arguments Formula
	 * MultiOperandNumericFunction's constructor has 2 parameters.
	 * the first parameter determines whether to evaluate boolean value. If it's
	 * true, evaluator will evaluate boolean value to double. TRUE to 1 and FALSE to 0.
	 * If it's false, boolean value is just ignored.
	 * The second parameter determines whether to evaluate blank value. 
	 * Blank value will be evaluated to 0.
	 */
	private static Function MY_SUBTOTAL = 
			new MultiOperandNumericFunction(false, false) {
		
		@Override
		protected double evaluate(double[] values) throws EvaluationException {
			double sum = 0;
			for (int i = 0 ; i < values.length ; i++){
				sum += values[i];
			}
			return sum;
		}
	}; 
	
	/**
	 * This formula depends on MultiOperandNumericFunction that evaluates all arguments to double.
	 * This formula sums all double values in cells.
	 * @param args
	 * @param srcCellRow row index of the cell containing the formula under evaluation
	 * @param srcCellCol column index of the cell containing the formula under evaluation
	 * @return
	 */
	public static ValueEval mySubtotal(ValueEval[] args, int srcCellRow, int srcCellCol){
		return MY_SUBTOTAL.evaluate(args, srcCellRow, srcCellCol); 
	}

	/**
	 * Advanced - Manually-Handled Arguments Formula. 
	 * This method demonstrates how to evaluate arguments manually. 
	 * The method implements a formula that concatenates all texts in cells 
	 * and ignores other types of value.
	 * 
	 * @param args the evaluated formula arguments
	 * @param srcCellRow unused
	 * @param srcCellCol unused
	 * @return calculated result, a subclass of ValueEval
	 */
	public static ValueEval chain(ValueEval[] args, int srcCellRow, int srcCellCol){

		List<StringEval> stringList = new LinkedList<StringEval>();
		for (int i = 0 ; i < args.length ; i++){
			//process an argument like A1:B2
			if (args[i] instanceof TwoDEval) {
				TwoDEval twoDEval = (TwoDEval) args[i];
				int width = twoDEval.getWidth();
				int height = twoDEval.getHeight();
				for (int rowIndex=0; rowIndex<height; rowIndex++) {
					for (int columnIndex=0; columnIndex<width; columnIndex++) {
						ValueEval ve = twoDEval.getValue(rowIndex, columnIndex);
						if (ve instanceof StringEval){
							stringList.add((StringEval)ve);
						}
					}
				}
				continue;
			}
			//process an argument like C18
			if (args[i] instanceof RefEval){
				ValueEval valueEval = ((RefEval)args[i]).getInnerValueEval();
				if (valueEval instanceof StringEval){
					stringList.add((StringEval)valueEval);
				}
				continue;
			}
			if (args[i] instanceof StringEval){
				stringList.add((StringEval)args[i]);
				continue;
			}

		}
		//chain all string value
		StringBuffer result = new StringBuffer();
		for (StringEval s: stringList){
			result.append(s.getStringValue());
		}

		//throw EvaluationException if encounter error conditions
		return new StringEval(result.toString());
	}

}
