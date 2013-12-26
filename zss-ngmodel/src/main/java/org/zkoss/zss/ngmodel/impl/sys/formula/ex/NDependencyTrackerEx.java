/* ELVariableEvaluator.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula.ex;

import org.zkoss.poi.ss.formula.DependencyTracker;
import org.zkoss.poi.ss.formula.OperationEvaluationContext;
import org.zkoss.poi.ss.formula.UserDefinedFunction;
import org.zkoss.poi.ss.formula.eval.NameEval;
import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.ptg.Ptg;

/**
 * An EL Variable evaluator through POI dependency tracker and ignore POI dependency tracking
 * @author Pao
 */
public class NDependencyTrackerEx implements DependencyTracker {

	@Override
	public ValueEval postProcessValueEval(OperationEvaluationContext ec, ValueEval opResult, boolean eval) {
		ValueEval opResultX = opResult;
		if(eval && opResult instanceof NameEval) { // resolve variable
			final String name = ((NameEval)opResult).getFunctionName();
			try {
				opResultX = UserDefinedFunction.instance.evaluate(new ValueEval[]{
						new NameEval(NFunctionResolverEx.EL_FUNCTION_KEY), new StringEval("${" + name + "}")},
						ec);
			} catch(NotImplementedException ex) {
				// ignore
				System.err.println(ex);
			}
		}
		return opResultX;
	}

	@Override
	public void addDependency(OperationEvaluationContext ec, Ptg[] ptgs) {
		// do nothing
	}

}
