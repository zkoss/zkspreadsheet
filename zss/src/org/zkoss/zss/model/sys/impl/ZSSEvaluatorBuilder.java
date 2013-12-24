/* ZSSEvaluatorBuilder.java

	Purpose:
		
	Description:
		
	History:
		Dec 24, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model.sys.impl;

import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.xssf.model.IndexedUDFFinder;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.XelContext;
import org.zkoss.xel.util.SimpleXelContext;
import org.zkoss.zss.formula.DefaultFunctionResolver;
import org.zkoss.zss.formula.FunctionResolver;
import org.zkoss.zss.formula.NoCacheClassifier;
import org.zkoss.zss.ngmodel.impl.sys.formula.EvaluatorBuilder;

/**
 * @author Pao
 */
public class ZSSEvaluatorBuilder implements EvaluatorBuilder {

	private FunctionMapper extraFuncMapper;

	public ZSSEvaluatorBuilder(FunctionMapper extraFuncMapper) {
		this.extraFuncMapper = extraFuncMapper;
	}

	@Override
	public WorkbookEvaluator createEvaluator(EvaluationWorkbook evalBook) {

		// create function resolver
		FunctionResolver resolver = (FunctionResolver)BookHelper.getLibraryInstance(FunctionResolver.CLASS);
		if(resolver == null) {
			resolver = new DefaultFunctionResolver();
		}

		// aggregate built-in functions and user defined functions
		UDFFinder udff = resolver.getUDFFinder();
		if(udff != null) {
			IndexedUDFFinder udffinder = (IndexedUDFFinder)evalBook.getUDFFinder();
			udffinder.insert(0, udff);
		}

		// create XEL context for resolving user defined function and EL variable
		JoinFunctionMapper functionMapper = new JoinFunctionMapper(resolver.getFunctionMapper());
		functionMapper.addFunctionMapper(extraFuncMapper);
		JoinVariableResolver variableResolver = new JoinVariableResolver();
		final XelContext ctx = new SimpleXelContext(variableResolver, functionMapper);
		ctx.setAttribute("zkoss.zss.CellType", Object.class);

		// create evaluator
		WorkbookEvaluator evaluator = new WorkbookEvaluator(evalBook, NoCacheClassifier.instance, TolerantUDFFinder.instance) {
			@Override
			public ValueEval evaluate(EvaluationCell srcCell) {
				// temporarily replace XEL context for resolving
				XelContext old = XelContextHolder.getXelContext();
				XelContextHolder.setXelContext(ctx);
				try {
					return super.evaluate(srcCell);
				} finally {
					XelContextHolder.setXelContext(old); // rollback XEL context
				}
			}

			@Override
			public ValueEval evaluate(int sheetIndex, String formula, boolean ignoreDereference) {
				// temporarily replace XEL context for resolving
				XelContext old = XelContextHolder.getXelContext();
				XelContextHolder.setXelContext(ctx);
				try {
					return super.evaluate(sheetIndex, formula, ignoreDereference);
				} finally {
					XelContextHolder.setXelContext(old); // rollback XEL context
				}
			}
		};

		return evaluator;
	}

}
