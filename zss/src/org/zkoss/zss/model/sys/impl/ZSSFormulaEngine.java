/* ZSSFormulaEngine.java

	Purpose:
		
	Description:
		
	History:
		Dec 26, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model.sys.impl;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.xssf.model.IndexedUDFFinder;
import org.zkoss.xel.Function;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelContext;
import org.zkoss.xel.XelException;
import org.zkoss.xel.util.SimpleXelContext;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Page;
import org.zkoss.zss.formula.DefaultFunctionResolver;
import org.zkoss.zss.formula.FunctionResolver;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.sys.formula.EvalBook;
import org.zkoss.zss.ngmodel.impl.sys.formula.FormulaEngineImpl;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;

/**
 * A implementation of Formula Engine with ZSS evaluation context.
 * Let formula engine support user defined function, EL variable evaluation ant so on.
 * @author Pao
 */
public class ZSSFormulaEngine extends FormulaEngineImpl implements FormulaEngine {

	private Map<EvaluationWorkbook, XelContext> xelContexts = new HashMap<EvaluationWorkbook, XelContext>();

	public ZSSFormulaEngine() {
	}

	@Override
	protected EvaluationResult evaluateFormula(FormulaExpression expr, FormulaEvaluationContext context, EvalBook evalBook, WorkbookEvaluator evaluator) throws FormulaParseException, Exception {

		// temporarily replace XEL context for resolving
		XelContext old = XelContextHolder.getXelContext();
		XelContext ctx = getXelContext(context, evalBook);
		XelContextHolder.setXelContext(ctx);
		try {
			return super.evaluateFormula(expr, context, evalBook, evaluator);
		} finally {
			XelContextHolder.setXelContext(old);
		}
	}

	private XelContext getXelContext(FormulaEvaluationContext context, EvaluationWorkbook evalBook) {
		XelContext xelContext = xelContexts.get(evalBook);
		if(xelContext == null) {

			// create function resolver
			FunctionResolver resolver = (FunctionResolver)BookHelper
					.getLibraryInstance(FunctionResolver.CLASS);
			if(resolver == null) {
				resolver = new DefaultFunctionResolver();
			}

			// aggregate built-in functions and user defined functions
			UDFFinder zkUDFF = resolver.getUDFFinder(); // ZK user defined function finder
			if(zkUDFF != null) {
				IndexedUDFFinder bookUDFF = (IndexedUDFFinder)evalBook.getUDFFinder(); // book contained built-in function finder
				bookUDFF.insert(0, zkUDFF);
			}

			// create XEL context for resolving user defined function and EL variable
			JoinFunctionMapper functionMapper = new JoinFunctionMapper(null);
			FunctionMapper userFunctionMapper = context.getFunctionMapper();
			if(userFunctionMapper != null) {
				functionMapper.addFunctionMapper(userFunctionMapper); // before ZSS function mapper
			}
			functionMapper.addFunctionMapper(resolver.getFunctionMapper()); // ZSS function mapper
			JoinVariableResolver variableResolver = new JoinVariableResolver();
			VariableResolver pageVariableResolver = context.getVariableResolver();
			if(pageVariableResolver != null) {
				variableResolver.addVariableResolver(pageVariableResolver);
			}
			xelContext = new SimpleXelContext(variableResolver, functionMapper);
			xelContext.setAttribute("zkoss.zss.CellType", Object.class);

			xelContexts.put(evalBook, xelContext);
		}
		return xelContext;
	}

	// FIXME zss 3.5, for testing now and should be modified
	@Override
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context) {
		context = new FormulaEvaluationContextWrapper(context);
		return super.evaluate(expr, context);
	}

	private static class FormulaEvaluationContextWrapper extends FormulaEvaluationContext {
		private FormulaEvaluationContext context;
		private FunctionMapper funcMapper = new PageFunctionMapper();

		public FormulaEvaluationContextWrapper(FormulaEvaluationContext context) {
			super((NBook)null);
			this.context = context;
		}

		public FunctionMapper getFunctionMapper() {
			return funcMapper;
		}

		public Locale getLocale() {
			return context.getLocale();
		}

		public void setLocale(Locale locale) {
			context.setLocale(locale);
		}

		public NBook getBook() {
			return context.getBook();
		}

		public NSheet getSheet() {
			return context.getSheet();
		}

		public NCell getCell() {
			return context.getCell();
		}

		public VariableResolver getVariableResolver() {
			return context.getVariableResolver();
		}

		public int hashCode() {
			return context.hashCode();
		}

		public boolean equals(Object obj) {
			return context.equals(obj);
		}

		public String toString() {
			return context.toString();
		}
	}

	private static class PageFunctionMapper implements FunctionMapper, Serializable {
		private static final long serialVersionUID = 1L;

		private Page getPage() {
			Execution execution = Executions.getCurrent();
			return execution != null ? execution.getDesktop().getFirstPage() : null;
		}

		public Function resolveFunction(String prefix, String name) throws XelException {
			final Page page = getPage();
			if(page != null) {
				final FunctionMapper mapper = page.getFunctionMapper();
				if(mapper != null) {
					return mapper.resolveFunction(prefix, name);
				}
			}
			return null;
		}
	}
}
