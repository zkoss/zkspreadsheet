package org.zkoss.zss.essential.advanced;

import java.util.LinkedList;
import java.util.List;

import org.zkoss.poi.ss.formula.TwoDEval;
import org.zkoss.poi.ss.formula.eval.EvaluationException;
import org.zkoss.poi.ss.formula.eval.RefEval;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.functions.MultiOperandNumericFunction;

/**
 * implement a formula that accepts multiple operations
 * @author Hawk
 *
 */
public class MyMultiOperandFormulas {

	// the first parameter determines whether to evaluate boolean value. If it's
	// true, evaluator will evaluate boolean value to number. TRUE to 1 and FALSE to 0.
	// If it's false, boolean value is just ignored.
	// The second parameter determines whether to evaluate blank value
	private static MultiOperandNumericFunction MY_SUBTOTAL = 
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
	 * This formula sums all double values in cells
	 * @param args
	 * @param srcCellRow
	 * @param srcCellCol
	 * @return
	 */
	public static ValueEval mySubtotal(ValueEval[] args, int srcCellRow, int srcCellCol){
		return MY_SUBTOTAL.evaluate(args, srcCellRow, srcCellCol); 
	}

	/**
	 * This formula concatenates all texts in cells and ignores other types of value
	 * @param args
	 * @param srcCellRow
	 * @param srcCellCol
	 * @return
	 */
	public static ValueEval chain(ValueEval[] args, int srcCellRow, int srcCellCol){
		
		//filter
		List<StringEval> stringList = new LinkedList<StringEval>();
		for (int i = 0 ; i < args.length ; i++){
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
			}
			if (args[i] instanceof RefEval){
				ValueEval valueEval = ((RefEval)args[i]).getInnerValueEval();
				if (valueEval instanceof StringEval){
					stringList.add((StringEval)valueEval);
				}
			}
			if (args[i] instanceof StringEval){
				stringList.add((StringEval)args[i]);
			}
		}
		StringBuffer result = new StringBuffer();
		for (StringEval s: stringList){
			result.append(s.getStringValue());
		}
		
		return new StringEval(result.toString());
	}
	
//	private static AggregateFunction MY_SUBTOTAL2 = new AggregateFunction() {
//		
//		@Override
//		protected double evaluate(double[] values) throws EvaluationException{
//			double sum = 0;
//			for (int i = 0 ; i < values.length ; i++){
//				sum += values[i];
//			}
//			return sum;
//		}
//	};
}
