/* ELEvaluate.java

	Purpose:
		
	Description:
		
	History:
		Apr 6, 2010 11:55:15 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula;

import java.util.Date;

import org.apache.poi.hssf.record.formula.eval.BoolEval;
import org.apache.poi.hssf.record.formula.eval.ErrorEval;
import org.apache.poi.hssf.record.formula.eval.EvaluationException;
import org.apache.poi.hssf.record.formula.eval.NumberEval;
import org.apache.poi.hssf.record.formula.eval.OperandResolver;
import org.apache.poi.hssf.record.formula.eval.StringEval;
import org.apache.poi.hssf.record.formula.eval.ValueEval;
import org.apache.poi.hssf.record.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.zkoss.xel.Expressions;
import org.zkoss.xel.XelContext;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.XelContextHolder;

/**
 * A function that do EL evaluation.
 * @author henrichen
 *
 */
public class ELEvaluate implements FreeRefFunction {

	public static final FreeRefFunction instance = new ELEvaluate();

	private ELEvaluate() {
		// enforce singleton
	}
	
	@Override
	public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
		if (args.length != 1) {
			return ErrorEval.NAME_INVALID;
		}
		try {
			final ValueEval ve = OperandResolver.getSingleValue(args[0], ec.getRowIndex(), ec.getColumnIndex());
			final String expression = OperandResolver.coerceValueToString(ve);
			final XelContext ctx = XelContextHolder.getXelContext();
			final Object obj = Expressions.evaluate(ctx, expression, (Class<?>) ctx.getAttribute("zkoss.zss.CellType"));
			if (obj == null) {
				return ErrorEval.NAME_INVALID;
			}
			if (obj instanceof String) {
				return new StringEval((String)obj);
			} else if (obj instanceof Number) {
				return new NumberEval(((Number)obj).doubleValue());
			} else if (obj instanceof Boolean) {
				return BoolEval.valueOf(((Boolean)obj).booleanValue());
			} else if (obj instanceof Date) {
				return new NumberEval(javaMillSecondToExcelDate(((Date)obj).getTime()));
			}
		} catch (EvaluationException e) {
			return e.getErrorEval();
		}


		return null;
	}

	private double javaMillSecondToExcelDate(long date) {
		return date / (1000 * 60 * 60 * 24);
	}
}
