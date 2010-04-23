/* ELEvalFunction.java

	Purpose:
		
	Description:
		
	History:
		Apr 19, 2010 2:23:06 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula;

import java.lang.reflect.InvocationTargetException;
import java.util.Date;

import org.apache.poi.hssf.record.formula.eval.BoolEval;
import org.apache.poi.hssf.record.formula.eval.ErrorEval;
import org.apache.poi.hssf.record.formula.eval.EvaluationException;
import org.apache.poi.hssf.record.formula.eval.NumberEval;
import org.apache.poi.hssf.record.formula.eval.OperandResolver;
import org.apache.poi.hssf.record.formula.eval.StringEval;
import org.apache.poi.hssf.record.formula.eval.ValueEval;
import org.apache.poi.hssf.record.formula.functions.Function;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.xel.XelContext;
import org.zkoss.zss.model.impl.XelContextHolder;

/**
 * This the default function that delegate the POI function call to EL tld function.
 * @author henrichen
 *
 */
public class ELEvalFunction implements Function {
	private final String _functionName;

	public ELEvalFunction(String name) {
		_functionName = name;
	}
	
	@Override
	public ValueEval evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		final XelContext ctx = XelContextHolder.getXelContext();
		if (ctx != null) {
			org.zkoss.xel.Function fn = ctx.getFunctionMapper().resolveFunction("", _functionName);
			if (fn != null) {
				try {
					return (ValueEval) fn.invoke(null, new Object[] {args, srcRowIndex, srcColumnIndex});
				} catch (Exception e) {
					if (e instanceof InvocationTargetException) {
						final Throwable te = ((InvocationTargetException)e).getTargetException();
						if (te instanceof EvaluationException) {
							return ((EvaluationException) te).getErrorEval();
						}
					} else {
						return ErrorEval.VALUE_INVALID;
					}
				}
			}
		}
		throw new NotImplementedException(_functionName);
	}
}
