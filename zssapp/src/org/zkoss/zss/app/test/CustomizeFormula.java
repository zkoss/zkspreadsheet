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
package org.zkoss.zss.app.test;

//import org.zkoss.poi.ss.formula.eval.EvaluationException;
//import org.zkoss.poi.ss.formula.eval.ValueEval;
//import org.zkoss.poi.ss.formula.functions.Function;
//import org.zkoss.poi.ss.formula.functions.MultiOperandNumericFunction;

/**
 * @author sam
 *
 */
public class CustomizeFormula {
	
//	public static final ValueEval foo(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
//		return FOO.evaluate(args, srcRowIndex, srcColumnIndex);
//	}
	
	public static final ValueEval bar(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return BAR.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	
//	public static final Function FOO = new NumericFunction() {
//
//		@Override
//		public double eval(ValueEval[] args, int srcRowIndex, int srcColumnIndex) throws EvaluationException {
//
//			final List ls = UtilFns.toList(args, srcRowIndex, srcColumnIndex);
//			if (ls.isEmpty()) {
//				throw new EvaluationException(ErrorEval.DIV_ZERO);
//			}
//			
//			for (int i = 0; i < ls.size(); i++) {
//				System.out.println(ls.get(i).toString());
//			}
//			
//			return 3.1415;
//		}
//	};
	
	public static final Function BAR = new MultiOperandNumericFunction(false, false) {

		@Override
		protected double evaluate(double[] values) throws EvaluationException {
			for (int i = 0; i < values.length; i++) {
				System.out.println(values[i]);
			}
			
			return 3.1415;
		}
	};
}
