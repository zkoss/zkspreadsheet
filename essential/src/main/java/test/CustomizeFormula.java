/* CustomizeFormula.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 29, 2012 3:29:49 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package test;

import org.zkoss.poi.ss.formula.eval.EvaluationException;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.functions.Function;
import org.zkoss.poi.ss.formula.functions.MultiOperandNumericFunction;
import org.zkoss.poi.ss.formula.functions.NumericFunction;

/**
 * @author sam
 *
 */
public class CustomizeFormula {
	
	public static final ValueEval foo(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FOO.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	
	public static final ValueEval bar(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return BAR.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	
	public static final Function FOO = new NumericFunction() {

		@Override
		public double eval(ValueEval[] args, int srcRowIndex, int srcColumnIndex) throws EvaluationException {

//			final List ls = UtilFns.toList(args, srcRowIndex, srcColumnIndex);
//			if (ls.isEmpty()) {
//				throw new EvaluationException(ErrorEval.DIV_ZERO);
//			}
			
			double sum;
			
			for(ValueEval v:args){
				
			}
			
			for (int i = 0; i < args.length; i++) {
				System.out.println(">>>FOO"+args[i]);
			}
			
			return 3.1415;
		}
	};
	
	public static final Function BAR = new MultiOperandNumericFunction(false, false) {

		@Override
		protected double evaluate(double[] values) throws EvaluationException {
			
			double sum = 0;
			
			for (int i = 0; i < values.length; i++) {
				sum+=values[i];
			}
			
			return sum;
		}
	};
}
