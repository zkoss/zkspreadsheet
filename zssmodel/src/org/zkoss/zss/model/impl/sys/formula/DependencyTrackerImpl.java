/* DependencyTrackerImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model.impl.sys.formula;

import org.zkoss.poi.ss.formula.DependencyTracker;
import org.zkoss.poi.ss.formula.OperationEvaluationContext;
import org.zkoss.poi.ss.formula.eval.ErrorEval;
import org.zkoss.poi.ss.formula.eval.NameEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.ptg.Ptg;

/**
 * A default dependency tracker.
 * @author Pao
 */
public class DependencyTrackerImpl implements DependencyTracker {

	@Override
	public ValueEval postProcessValueEval(OperationEvaluationContext ec, ValueEval opResult, boolean eval) {
		// resolving name to be invalid 
		if(eval && opResult instanceof NameEval) {
			return ErrorEval.NAME_INVALID;
		}
		return opResult;
	}

	@Override
	public void addDependency(OperationEvaluationContext ec, Ptg[] ptgs) {
		// do nothing, we don't need POI dependency tracking
	}

}
