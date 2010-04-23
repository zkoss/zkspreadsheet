/* FinanceFunction.java

	Purpose:
		
	Description:
		
	History:
		Apr 19, 2010 4:34:26 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula.fn;

import org.apache.poi.hssf.record.formula.eval.ErrorEval;
import org.apache.poi.hssf.record.formula.eval.EvaluationException;
import org.apache.poi.hssf.record.formula.eval.ValueEval;
import org.apache.poi.hssf.record.formula.functions.FinanceFunction;
import org.apache.poi.hssf.record.formula.functions.Function;

/**
 * Excel financial functions.
 * @author henrichen
 */
public class FinancialFns  {
	private FinancialFns() {
	}
	
	public static final ValueEval db(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FinanceFunctionImpl.DB.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	public static final ValueEval ddb(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FinanceFunctionImpl.DDB.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	public static final ValueEval ipmt(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FinanceFunctionImpl.IPMT.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	public static final ValueEval ppmt(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FinanceFunctionImpl.PPMT.evaluate(args, srcRowIndex, srcColumnIndex);
	}
	public static final ValueEval sln(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FinanceFunctionImpl.SLN.evaluate(args, srcRowIndex, srcColumnIndex);
	};
	public static final ValueEval syd(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		return FinanceFunctionImpl.SYD.evaluate(args, srcRowIndex, srcColumnIndex);
	};
}
