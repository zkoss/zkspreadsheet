/* ZSSFormulaEngine.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model.sys.impl;

import org.zkoss.xel.XelContext;
import org.zkoss.zss.ngmodel.impl.sys.formula.FormulaEngineImpl;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;

/**
 * A implementation of Formula Engine with XEL context.
 * Let formula engine support user defined function, EL variable evaluation ant so on.
 * @author Pao
 */
public class ZSSFormulaEngine extends FormulaEngineImpl implements FormulaEngine {

	public ZSSFormulaEngine() {
	}

	protected Object getXelContext() {
		return XelContextHolder.getXelContext();
	}

	protected void setXelContext(Object ctx) {
		XelContextHolder.setXelContext((XelContext)ctx);
	}
}
