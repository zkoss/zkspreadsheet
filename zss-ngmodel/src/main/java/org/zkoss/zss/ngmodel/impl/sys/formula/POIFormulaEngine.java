/* ZPOIEngine.java

	Purpose:
		
	Description:
		
	History:
		Dec 10, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import java.io.Serializable;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.zkoss.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.eval.ErrorEval;
import org.zkoss.poi.ss.formula.eval.NumberEval;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.eval.ValuesEval;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.impl.BookSeriesAdv;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;

/**
 * A formula engine implemented by ZPOI
 * @author Pao
 */
public class POIFormulaEngine implements FormulaEngine {
	private final static Logger logger = Logger.getLogger(POIFormulaEngine.class.getName());

	@Override
	public FormulaExpression parse(String formula, FormulaParseContext context) {
		FormulaExpression expr = null;
		try {
			// adapt and parse
			NBook book = context.getBook();
			int sheetIndex = book.getSheetIndex(context.getSheet());
			EvalBook evalBook = new EvalBook(book);
			evalBook.getFormulaTokens(sheetIndex, formula);
			expr = new FormulaExpressionImpl(formula, true);
		} catch(FormulaParseException e) {
			logger.log(Level.INFO, e.getMessage(), e);
			expr = new FormulaExpressionImpl(formula, false);
		}
		return expr;
	}

	@Override
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context) {

		EvaluationResult result = null;
		try {
			NBook book = context.getBook();
			WorkbookEvaluator evaluator = null;

			// book series
			BookSeriesAdv bookSeries = (BookSeriesAdv)book.getBookSeries();
			NBook[] books = bookSeries.getBooks().toArray(new NBook[0]);
			String[] bookNames = new String[books.length];
			WorkbookEvaluator[] evaluators = new WorkbookEvaluator[books.length];
			for(int i = 0; i < books.length; ++i) {
				bookNames[i] = books[i].getBookName();
				EvalBook eb = new EvalBook(books[i]);
				evaluators[i] = new WorkbookEvaluator(eb, null, null);
				if(context.getBook() == books[i]) {
					evaluator = evaluators[i];
				}
			}
			CollaboratingWorkbooksEnvironment.setup(bookNames, evaluators);
			if(evaluator == null) { // just in case
				return new EvaluationResultImpl(ResultType.ERROR, "The book isn't in the book series.");
			}

			// FIXME dependency tracker
			// ((BookSeriesAdv)book.getBookSeries()).getDependencyTable();
			// DependencyTracker tracker = new DependencyTracker();
			// evaluator.setDependencyTracker(new DependencyTrackerAdapter("book1", tracker));

			// evaluation formula directly
			int currentSheetIndex = book.getSheetIndex(context.getSheet());
			ValueEval value = evaluator.evaluate(currentSheetIndex, expr.getFormulaString(), false);

			// convert to result
			if(value instanceof ValuesEval) { // FIXME check these code
				value = ((ValuesEval)value).getValueEvals()[0];
			}
			if(value instanceof ErrorEval) {
				int code = ((ErrorEval)value).getErrorCode();
				result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue((byte)code));
			} else if(value instanceof StringEval) {
				result = new EvaluationResultImpl(ResultType.SUCCESS, ((StringEval)value).getStringValue());
			} else if(value instanceof NumberEval) {
				result = new EvaluationResultImpl(ResultType.SUCCESS, ((NumberEval)value).getNumberValue());
			} else {
				throw new Exception("no matched"); // FIXME
			}
		} catch(Exception e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
			result = new EvaluationResultImpl(ResultType.ERROR, null);
		}
		return result;
	}

	private static class FormulaExpressionImpl implements FormulaExpression, Serializable {
		private static final long serialVersionUID = -8532826169759927711L;
		private String formula;
		private boolean valid;

		public FormulaExpressionImpl(String formula, boolean valid) {
			this.formula = formula;
			this.valid = valid;
		}

		@Override
		public boolean hasError() {
			return !valid;
		}

		@Override
		public String getFormulaString() {
			return formula;
		}

		@Override
		public String reformSheetNameChanged(String oldName, String newName) {
			// TODO
			return null;
		}

	}

	private static class EvaluationResultImpl implements EvaluationResult {

		private ResultType type;
		private Object value;

		public EvaluationResultImpl(ResultType type, Object value) {
			this.type = type;
			this.value = value;
		}

		@Override
		public ResultType getType() {
			return type;
		}

		@Override
		public Object getValue() {
			return value;
		}

	}

}
