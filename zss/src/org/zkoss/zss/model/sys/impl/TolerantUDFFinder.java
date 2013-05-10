package org.zkoss.zss.model.sys.impl;

import org.zkoss.poi.ss.formula.OperationEvaluationContext;
import org.zkoss.poi.ss.formula.eval.ErrorEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.functions.FreeRefFunction;
import org.zkoss.poi.ss.formula.udf.UDFFinder;

/**
 * Return an user function that always evaluate to
 * {@code ErrorEval#NAME_INVALID} to prevent poi throws NotImplementedException
 * 
 * @author dennis
 * 
 */
public class TolerantUDFFinder implements UDFFinder {

	/* package */static final TolerantUDFFinder instance = new TolerantUDFFinder();
	private static final FreeRefFunction toerantFunction = new FreeRefFunction() {
		public ValueEval evaluate(ValueEval[] args,
				OperationEvaluationContext ec) {
			return ErrorEval.NAME_INVALID;
		}
	};

	private TolerantUDFFinder() {

	}

	@Override
	public FreeRefFunction findFunction(String name) {
		return toerantFunction;
	}

}
