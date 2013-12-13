/* ZPOIEngine.java

	Purpose:
		
	Description:
		
	History:
		Dec 10, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import java.io.Serializable;
import java.lang.reflect.Proxy;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.zkoss.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.eval.ErrorEval;
import org.zkoss.poi.ss.formula.eval.NumberEval;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.eval.ValuesEval;
import org.zkoss.poi.ss.formula.ptg.Area3DPtg;
import org.zkoss.poi.ss.formula.ptg.AreaPtg;
import org.zkoss.poi.ss.formula.ptg.FuncPtg;
import org.zkoss.poi.ss.formula.ptg.NamePtg;
import org.zkoss.poi.ss.formula.ptg.NameXPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.ptg.Ref3DPtg;
import org.zkoss.poi.ss.formula.ptg.RefPtg;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.BookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTableAdv;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
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
			EvaluationWorkbook evalBook = new EvalBook(book);
			evalBook = (EvaluationWorkbook)Proxy.newProxyInstance(ClassLoader.getSystemClassLoader(),
					new Class<?>[]{EvaluationWorkbook.class}, new LogHandler(evalBook));
			Ptg[] tokens = evalBook.getFormulaTokens(sheetIndex, formula);

			// dependency tracking
			BookSeriesAdv series = (BookSeriesAdv)book.getBookSeries();
			DependencyTableAdv table = (DependencyTableAdv)series.getDependencyTable();
			Ref dependant = context.getDependent();
			for(Ptg ptg : tokens) {
				Ref precedent = toDenpendRef(context, evalBook, ptg);
				if(precedent != null) {
					table.add(dependant, precedent);
				}
			}

			// create result
			expr = new FormulaExpressionImpl(formula, true);
		} catch(FormulaParseException e) {
			logger.log(Level.INFO, e.getMessage(), e);
			expr = new FormulaExpressionImpl(formula, false);
		}
		return expr;
	}

	private Ref toDenpendRef(FormulaParseContext ctx, EvaluationWorkbook evalBook, Ptg ptg) {
		try {
			NSheet sheet = ctx.getSheet();

			if(ptg instanceof NamePtg) {
				// TODO
				// NamePtg namePtg = (NamePtg)ptg;
				// EvaluationName nameRecord = ec.getWorkbook().getName(namePtg);
				// return new NameEval(nameRecord.getNameText());
				// new ObjectRefImpl(name, objectId) // how to create a ref before name existed
			} else if(ptg instanceof NameXPtg) {
				// TODO name in external book
				// return ec.getNameXEval(((NameXPtg)ptg));
			} else if(ptg instanceof Ref3DPtg) {
				Ref3DPtg aptg = (Ref3DPtg)ptg;
				String bookName = ctx.getBook().getBookName();
				int sheetIndex1 = evalBook.convertFromExternSheetIndex(aptg.getExternSheetIndex());
				int sheetIndex2 = evalBook.convertLastIndexFromExternSheetIndex(aptg.getExternSheetIndex());
				if(sheetIndex1 == sheetIndex2) {
					return new RefImpl(bookName, evalBook.getSheetName(sheetIndex1), aptg.getRow(),
							aptg.getColumn());
				} else {
					return new RefImpl(bookName, evalBook.getSheetName(sheetIndex1),
							evalBook.getSheetName(sheetIndex2), aptg.getRow(), aptg.getColumn());
				}
			} else if(ptg instanceof Area3DPtg) {
				Area3DPtg aptg = (Area3DPtg)ptg;
				String bookName = ctx.getBook().getBookName();
				int sheetIndex1 = evalBook.convertFromExternSheetIndex(aptg.getExternSheetIndex());
				int sheetIndex2 = evalBook.convertLastIndexFromExternSheetIndex(aptg.getExternSheetIndex());
				if(sheetIndex1 == sheetIndex2) {
					return new RefImpl(bookName, evalBook.getSheetName(sheetIndex1), aptg.getFirstRow(),
							aptg.getFirstColumn(), aptg.getLastRow(), aptg.getLastColumn());
				} else {
					return new RefImpl(bookName, evalBook.getSheetName(sheetIndex1),
							evalBook.getSheetName(sheetIndex2), aptg.getFirstRow(), aptg.getFirstColumn(),
							aptg.getLastRow(), aptg.getLastColumn());
				}
			} else if(ptg instanceof RefPtg) {
				RefPtg rptg = (RefPtg)ptg;
				String bookName = sheet.getBook().getBookName();
				String sheetName = sheet.getSheetName();
				return new RefImpl(bookName, sheetName, rptg.getRow(), rptg.getColumn());
			} else if(ptg instanceof AreaPtg) {
				AreaPtg aptg = (AreaPtg)ptg;
				String sheetName = sheet.getSheetName();
				String bookName = sheet.getBook().getBookName();
				return new RefImpl(bookName, sheetName, aptg.getFirstRow(), aptg.getFirstColumn(),
						aptg.getLastRow(), aptg.getLastColumn());
			} else if(ptg instanceof FuncPtg) {
				// TODO consider function-type dependency
			}
		} catch(Exception e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
		}
		return null;
	}

	@Override
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context) {

		EvaluationResult result = null;
		try {
			NBook book = context.getBook();
			EvalBook evalBook = null;
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
					evalBook = eb;
					evaluator = evaluators[i];
				}
			}
			CollaboratingWorkbooksEnvironment.setup(bookNames, evaluators);
			if(evalBook == null || evaluator == null) { // just in case
				return new EvaluationResultImpl(ResultType.ERROR, "The book isn't in the book series.");
			}

			// evaluation formula
			ValueEval value = null;
			int currentSheetIndex = book.getSheetIndex(context.getSheet());
			if(context.getCell().isNull()) {
				// evaluation formula directly
				value = evaluator.evaluate(currentSheetIndex, expr.getFormulaString(), false);
			} else {
				NCell cell = context.getCell();
				EvaluationCell evalCell = evalBook.getSheet(currentSheetIndex).getCell(cell.getRowIndex(),
						cell.getColumnIndex());
				value = evaluator.evaluate(evalCell);
			}

			// convert to result
			if(value instanceof ErrorEval) {
				int code = ((ErrorEval)value).getErrorCode();
				result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue((byte)code));
			} else if(value instanceof StringEval) {
				result = new EvaluationResultImpl(ResultType.SUCCESS, ((StringEval)value).getStringValue());
			} else if(value instanceof NumberEval) {
				result = new EvaluationResultImpl(ResultType.SUCCESS, ((NumberEval)value).getNumberValue());
			} else if(value instanceof ValuesEval) {
				ValueEval[] values = ((ValuesEval)value).getValueEvals();
				Object[] array = new Object[values.length];
				for(int i = 0; i < values.length; ++i) {
					if(value instanceof StringEval) {
						array[i] = ((StringEval)values[i]).getStringValue();
					} else if(value instanceof NumberEval) {
						array[i] = ((NumberEval)values[i]).getNumberValue();
					} else {
						throw new Exception("no matched type: " + array[i]); // FIXME
					}
				}
				return new EvaluationResultImpl(ResultType.SUCCESS, array);
			} else {
				throw new Exception("no matched type: " + value); // FIXME
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

		@Override
		public boolean isRefersTo() {
			return false;
		}

		@Override
		public String getRefersToSheetName() {
			return null;
		}

		@Override
		public CellRegion getRefersToCellRegion() {
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
