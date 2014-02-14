/* ZPOIEngine.java

	Purpose:
		
	Description:
		
	History:
		Dec 10, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.zkoss.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.zkoss.poi.ss.formula.DependencyTracker;
import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationSheet;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.EvaluationWorkbook.ExternalSheet;
import org.zkoss.poi.ss.formula.ExternSheetReferenceToken;
import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.IStabilityClassifier;
import org.zkoss.poi.ss.formula.PtgShifter;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.eval.AreaEval;
import org.zkoss.poi.ss.formula.eval.BlankEval;
import org.zkoss.poi.ss.formula.eval.BoolEval;
import org.zkoss.poi.ss.formula.eval.ErrorEval;
import org.zkoss.poi.ss.formula.eval.EvaluationException;
import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.poi.ss.formula.eval.NumberEval;
import org.zkoss.poi.ss.formula.eval.RefEval;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.eval.ValuesEval;
import org.zkoss.poi.ss.formula.ptg.Area3DPtg;
import org.zkoss.poi.ss.formula.ptg.AreaPtg;
import org.zkoss.poi.ss.formula.ptg.AreaPtgBase;
import org.zkoss.poi.ss.formula.ptg.FuncPtg;
import org.zkoss.poi.ss.formula.ptg.NamePtg;
import org.zkoss.poi.ss.formula.ptg.NameXPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.ptg.Ref3DPtg;
import org.zkoss.poi.ss.formula.ptg.RefPtg;
import org.zkoss.poi.ss.formula.ptg.RefPtgBase;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.xssf.model.IndexedUDFFinder;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelContext;
import org.zkoss.xel.util.SimpleXelContext;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.SheetRegion;
import org.zkoss.zss.ngmodel.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.NameRefImpl;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.impl.sys.DependencyTableAdv;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaClearContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.sys.formula.NFunctionResolver;
import org.zkoss.zss.ngmodel.sys.formula.NFunctionResolverFactory;

/**
 * A formula engine implemented by ZPOI
 * @author Pao
 */
public class FormulaEngineImpl implements FormulaEngine {

	public final static String KEY_EVALUATORS = "$ZSS_EVALUATORS$";
	public static boolean EE_EDITION = true; // FIXME zss 3.5

	private final static Logger logger = Logger.getLogger(FormulaEngineImpl.class.getName());

	private Map<EvaluationWorkbook, XelContext> xelContexts = new HashMap<EvaluationWorkbook, XelContext>();
	
	// for POI formula evaluator
	protected final static IStabilityClassifier noCacheClassifier = new IStabilityClassifier() {
		public boolean isCellFinal(int sheetIndex, int rowIndex, int columnIndex) {
			return true;
		}
	};

	@Override
	public FormulaExpression parse(String formula, FormulaParseContext context) {
		FormulaExpression expr = null;
		try {
			// adapt and parse
			NBook book = context.getBook();
			ParsingBook parsingBook = new ParsingBook(book);
			int sheetIndex = parsingBook.getExternalSheetIndex(null, context.getSheet().getSheetName());
			Ptg[] tokens = FormulaParser.parse(formula, parsingBook, FormulaType.CELL, sheetIndex); // current sheet index in parsing is always 0

			// dependency tracking if necessary
			Ref dependant = context.getDependent();
			if(dependant != null) {
				AbstractBookSeriesAdv series = (AbstractBookSeriesAdv)book.getBookSeries();
				DependencyTableAdv table = (DependencyTableAdv)series.getDependencyTable();
				for(Ptg ptg : tokens) {
					Ref precedent = toDenpendRef(context, parsingBook, ptg);
					if(precedent != null) {
						table.add(dependant, precedent);
					}
				}
			}

			// render formula, detect region and create result
			String renderedFormula = renderFormula(parsingBook, formula, tokens);
			Ref singleRef = tokens.length == 1 ? toDenpendRef(context, parsingBook, tokens[0]) : null;
			expr = new FormulaExpressionImpl(renderedFormula, singleRef, false);
		} catch(FormulaParseException e) {
			logger.log(Level.INFO, e.getMessage());
			expr = new FormulaExpressionImpl(formula, null, true);
		}
		return expr;
	}
	
	protected String renderFormula(ParsingBook parsingBook, String formula, Ptg[] tokens) {
		boolean hit = false; // only render if necessary
		for(Ptg token : tokens) {
			// check it is a external book reference or not
			if(token instanceof ExternSheetReferenceToken) {
				ExternSheetReferenceToken externalRef = (ExternSheetReferenceToken)token;
				ExternalSheet externalSheet = parsingBook.getExternalSheet(externalRef.getExternSheetIndex());
				if(externalSheet != null) {
					hit = true;
					break;
				}
			}
		}
		return hit ? FormulaRenderer.toFormulaString(parsingBook, tokens) : formula;
	}

	protected Ref toDenpendRef(FormulaParseContext ctx, ParsingBook parsingBook, Ptg ptg) {
		try {
			NSheet sheet = ctx.getSheet();

			if(ptg instanceof NamePtg) {	// name range name
				NamePtg namePtg = (NamePtg)ptg;
				// use current book, we don't refer to other book's defined name
				String bookName = sheet.getBook().getBookName(); 
				String name = parsingBook.getNameText(namePtg);
				return new NameRefImpl(bookName, null, name); // assume name is book-scope
			} else if(ptg instanceof NameXPtg) { // user defined function name
				// TODO consider function-type dependency
			} else if(ptg instanceof FuncPtg) {
				// TODO consider function-type dependency
			} else if(ptg instanceof Ref3DPtg) {
				Ref3DPtg rptg = (Ref3DPtg)ptg;
				// might be internal or external book reference
				ExternalSheet es = parsingBook.getAnyExternalSheet(rptg.getExternSheetIndex());
				String bookName = es.getWorkbookName() != null ? es.getWorkbookName() : sheet.getBook().getBookName();
				String sheetName = es.getSheetName();
				String lastSheetName = es.getLastSheetName().equals(sheetName) ? null : es.getLastSheetName();
				return new RefImpl(bookName, sheetName, lastSheetName, rptg.getRow(), rptg.getColumn());
			} else if(ptg instanceof Area3DPtg) {
				Area3DPtg aptg = (Area3DPtg)ptg;
				// might be internal or external book reference
				ExternalSheet es = parsingBook.getAnyExternalSheet(aptg.getExternSheetIndex());
				String bookName = es.getWorkbookName() != null ? es.getWorkbookName() : sheet.getBook().getBookName();
				String sheetName = es.getSheetName();
				String lastSheetName = es.getLastSheetName().equals(sheetName) ? null : es.getLastSheetName();
				return new RefImpl(bookName, sheetName, lastSheetName, aptg.getFirstRow(),
						aptg.getFirstColumn(), aptg.getLastRow(), aptg.getLastColumn());
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
			}
		} catch(Exception e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
		}
		return null;
	}

	@Override
	@SuppressWarnings("unchecked")
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context) {

		// by pass if expression is invalid format
		if(expr.hasError()) {
			return new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_FORMULA));
		}

		EvaluationResult result = null;
		try {

			// get evaluation context from book series
			NBook book = context.getBook();
			AbstractBookSeriesAdv bookSeries = (AbstractBookSeriesAdv)book.getBookSeries();
			Map<String, EvalContext> evalCtxMap = (Map<String, EvalContext>)bookSeries
					.getAttribute(KEY_EVALUATORS);

			// get evaluation context or create new one if not existed
			if(evalCtxMap == null) {
				evalCtxMap = new LinkedHashMap<String, FormulaEngineImpl.EvalContext>();
				List<String> bookNames = new ArrayList<String>();
				List<WorkbookEvaluator> evaluators = new ArrayList<WorkbookEvaluator>();
				for(NBook nb : bookSeries.getBooks()) {
					String bookName = nb.getBookName();
					EvalBook evalBook = new EvalBook(nb);
					WorkbookEvaluator we = new WorkbookEvaluator(evalBook, noCacheClassifier, null);
					bookNames.add(bookName);
					evaluators.add(we);
					evalCtxMap.put(bookName, new EvalContext(evalBook, we));

					// aggregate built-in functions and user defined functions
					NFunctionResolver resolver = NFunctionResolverFactory.getFunctionResolver();
					UDFFinder zkUDFF = resolver.getUDFFinder(); // ZK user defined function finder
					if(zkUDFF != null) {
						IndexedUDFFinder bookUDFF = (IndexedUDFFinder)evalBook.getUDFFinder(); // book contained built-in function finder
						bookUDFF.insert(0, zkUDFF);
					}
				}
				CollaboratingWorkbooksEnvironment.setup(bookNames.toArray(new String[0]),
						evaluators.toArray(new WorkbookEvaluator[0]));
				bookSeries.setAttribute(KEY_EVALUATORS, evalCtxMap);
			}
			// check again
			EvalContext ctx = evalCtxMap.get(book.getBookName());
			if(ctx == null) { // just in case
				throw new IllegalStateException("The book isn't in the book series.");
			}
			EvalBook evalBook = ctx.getBook();
			WorkbookEvaluator evaluator = ctx.getEvaluator();

			// evaluation formula
			// for resolving, temporarily replace current XEL context 
			Object oldXelCtx = getXelContext();
			XelContext xelCtx = getXelContextForResolving(context, evalBook, evaluator);
			setXelContext(xelCtx);
			try {
				result = evaluateFormula(expr, context, evalBook, evaluator);
			} finally {
				setXelContext(oldXelCtx);
			}

		} catch(NotImplementedException e) {
			logger.log(Level.INFO, e.getMessage() + " when eval " + expr.getFormulaString());
			result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_NAME));
		} catch(FormulaParseException e) {
			// we skip evaluation if formula has parsing error
			// so if still occurring formula parsing exception, it should be a bug 
			logger.log(Level.SEVERE, e.getMessage() + " when eval " + expr.getFormulaString());
			result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_FORMULA));
		} catch(Exception e) {
			logger.log(Level.SEVERE, e.getMessage() + " when eval " + expr.getFormulaString(), e);
			result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_FORMULA));
		}
		return result;
	}

	protected EvaluationResult evaluateFormula(FormulaExpression expr, FormulaEvaluationContext context, EvalBook evalBook, WorkbookEvaluator evaluator) throws FormulaParseException, Exception {

		// do evaluate
		NBook book = context.getBook();
		int currentSheetIndex = book.getSheetIndex(context.getSheet());
		NCell cell = context.getCell();
		ValueEval value = null;
		if(cell == null || cell.isNull()) {
			// evaluation formula directly
			value = evaluator.evaluate(currentSheetIndex, expr.getFormulaString(), true);
		} else {
			EvaluationCell evalCell = evalBook.getSheet(currentSheetIndex).getCell(cell.getRowIndex(),
					cell.getColumnIndex());
			value = evaluator.evaluate(evalCell);
		}

		// convert to result
		if(value instanceof ErrorEval) {
			int code = ((ErrorEval)value).getErrorCode();
			return new EvaluationResultImpl(ResultType.ERROR, new ErrorValue((byte)code));
		} else if(value instanceof BlankEval) {
			return new EvaluationResultImpl(ResultType.SUCCESS, "");
		} else if(value instanceof StringEval) {
			return new EvaluationResultImpl(ResultType.SUCCESS, ((StringEval)value).getStringValue());
		} else if(value instanceof NumberEval) {
			return new EvaluationResultImpl(ResultType.SUCCESS, ((NumberEval)value).getNumberValue());
		} else if(value instanceof BoolEval) {
			return new EvaluationResultImpl(ResultType.SUCCESS, ((BoolEval)value).getBooleanValue());
		} else if(value instanceof ValuesEval) {
			ValueEval[] values = ((ValuesEval)value).getValueEvals();
			Object[] array = new Object[values.length];
			for(int i = 0; i < values.length; ++i) {
				array[i] = getResolvedValue(values[i]);
			}
			return new EvaluationResultImpl(ResultType.SUCCESS, array);
		} else if(value instanceof AreaEval) {
			// covert all values into an array
			List<Object> list = new ArrayList<Object>();
			AreaEval area = (AreaEval)value;
			for(int r = 0; r < area.getHeight(); ++r) {
				for(int c = 0; c < area.getWidth(); ++c) {
					ValueEval v = area.getValue(r, c);
					list.add(getResolvedValue(v));
				}
			}
			return new EvaluationResultImpl(ResultType.SUCCESS, list);
		} else if(value instanceof RefEval) {
			ValueEval ve = ((RefEval)value).getInnerValueEval();
			Object v = getResolvedValue(ve);
			return new EvaluationResultImpl(ResultType.SUCCESS, v);
		} else {
			throw new Exception("no matched type: " + value); // FIXME
		}
	}

	protected Object getResolvedValue(ValueEval value) throws EvaluationException {
		if(value instanceof StringEval) {
			return ((StringEval)value).getStringValue();
		} else if(value instanceof NumberEval) {
			return ((NumberEval)value).getNumberValue();
		} else if(value instanceof BlankEval) {
			return "";
		} else {
			throw new EvaluationException(null, "no matched type: " + value); // FIXME
		}
	}

	protected Object getXelContext() {
		return null; // do nothing
	}

	protected void setXelContext(Object ctx) {
		// do nothing
	}

	private XelContext getXelContextForResolving(FormulaEvaluationContext context, EvaluationWorkbook evalBook, WorkbookEvaluator evaluator) {
		XelContext xelContext = xelContexts.get(evalBook);
		if(xelContext == null) {

			// create function resolver
			NFunctionResolver resolver = NFunctionResolverFactory.getFunctionResolver();

			// apply POI dependency tracker for defined name resolving
			DependencyTracker tracker = resolver.getDependencyTracker();
			if(tracker != null) {
				evaluator.setDependencyTracker(tracker);
			}
			
			// collect all function mappers 
			NJoinFunctionMapper functionMapper = new NJoinFunctionMapper(null);
			FunctionMapper extraFunctionMapper = context.getFunctionMapper();
			if(extraFunctionMapper != null) {
				functionMapper.addFunctionMapper(extraFunctionMapper); // must before ZSS function mapper
			}
			FunctionMapper zssFuncMapper = resolver.getFunctionMapper();
			if(zssFuncMapper != null) {
				functionMapper.addFunctionMapper(zssFuncMapper);
			}
			
			// collect all variable resolvers
			NJoinVariableResolver variableResolver = new NJoinVariableResolver();
			VariableResolver extraVariableResolver = context.getVariableResolver();
			if(extraVariableResolver != null) {
				variableResolver.addVariableResolver(extraVariableResolver);
			}
			
			// create XEL context
			xelContext = new SimpleXelContext(variableResolver, functionMapper);
			xelContext.setAttribute("zkoss.zss.CellType", Object.class);
			xelContexts.put(evalBook, xelContext);
		}
		return xelContext;
	}

	@Override
	@SuppressWarnings("unchecked")
	public void clearCache(FormulaClearContext context) {
		try {
			NBook book = context.getBook();
			NSheet sheet = context.getSheet();
			NCell cell = context.getCell();

			// take evaluators from book series
			AbstractBookSeriesAdv bookSeries = (AbstractBookSeriesAdv)book.getBookSeries();
			Map<String, EvalContext> map = (Map<String, EvalContext>)bookSeries.getAttribute(KEY_EVALUATORS);

			// do nothing if not existed
			if(map == null) {
				return;
			}

			// clean cache and target is a cell
			if(cell != null && !cell.isNull()) {

				// do nothing if not existed
				EvalContext ctx = map.get(book.getBookName());
				if(ctx == null) {
					logger.log(Level.WARNING, "clear a non-existed book? >> " + book.getBookName());
					return;
				}

				// notify POI formula evaluator one of cell has been updated
				String sheetName = sheet.getSheetName();
				EvalBook evalBook = ctx.getBook();
				EvaluationSheet evalSheet = evalBook.getSheet(evalBook.getSheetIndex(sheetName));
				EvaluationCell evalCell = evalSheet.getCell(cell.getRowIndex(), cell.getColumnIndex());
				WorkbookEvaluator evaluator = ctx.getEvaluator();
				evaluator.notifyUpdateCell(evalCell);
			} else {
				// no cell indicates clearing all cache
				bookSeries.setAttribute(KEY_EVALUATORS, null);
				map.clear(); // just in case
			}
		} catch(Exception e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
		}
	}

	protected static class FormulaExpressionImpl implements FormulaExpression, Serializable {
		private static final long serialVersionUID = -8532826169759927711L;
		private String formula;
		private Ref ref;
		private boolean error;

		/**
		 * @param ref resolved reference if formula has only one parsed token
		 */
		public FormulaExpressionImpl(String formula, Ref ref, boolean error) {
			this.formula = formula;
			this.ref = ref;
			this.error = error;
		}

		@Override
		public boolean hasError() {
			return error;
		}

		@Override
		public String getFormulaString() {
			return formula;
		}

		@Override
		public boolean isRefersTo() {
			return ref != null && (ref.getType() == RefType.AREA || ref.getType() == RefType.CELL);
		}
		
		@Override
		public String getRefersToBookName() {
			return isRefersTo() ? ref.getBookName() : null;
		}
		
		@Override
		public String getRefersToSheetName() {
			return isRefersTo() ? ref.getSheetName() : null;
		}
		
		@Override
		public String getRefersToLastSheetName() {
			return isRefersTo() ? ref.getLastSheetName() : null;
		}

		@Override
		public CellRegion getRefersToCellRegion() {
			return isRefersTo() ? new CellRegion(ref.getRow(), ref.getColumn(), ref.getLastRow(),
					ref.getLastColumn()) : null;
		}

	}

	protected static class EvaluationResultImpl implements EvaluationResult {

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

	protected static class EvalContext {
		private EvalBook book;
		private WorkbookEvaluator evaluator;

		public EvalContext(EvalBook book, WorkbookEvaluator evaluator) {
			this.book = book;
			this.evaluator = evaluator;
		}

		public EvalBook getBook() {
			return book;
		}

		public WorkbookEvaluator getEvaluator() {
			return evaluator;
		}

		@Override
		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + ((book == null) ? 0 : book.hashCode());
			result = prime * result + ((evaluator == null) ? 0 : evaluator.hashCode());
			return result;
		}

		@Override
		public boolean equals(Object obj) {
			if(this == obj)
				return true;
			if(obj == null)
				return false;
			if(getClass() != obj.getClass())
				return false;
			EvalContext other = (EvalContext)obj;
			if(book == null) {
				if(other.book != null)
					return false;
			} else if(!book.equals(other.book))
				return false;
			if(evaluator == null) {
				if(other.evaluator != null)
					return false;
			} else if(!evaluator.equals(other.evaluator))
				return false;
			return true;
		}

	}

	@Override
	public FormulaExpression renameSheet(String formula, String oldName, String newName, FormulaParseContext context) {
		// TODO zss 3.5
		return new FormulaExpressionImpl(formula, null, true);
	}
	
	@Override
	public FormulaExpression move(String formula, SheetRegion region, int rowOffset, int columnOffset, FormulaParseContext context) {
		FormulaExpression expr = null;
		try {
			// adapt and parse
			ParsingBook parsingBook = new ParsingBook(context.getBook());
			String bookName = context.getBook().getBookName();
			String sheetName = context.getSheet().getSheetName();
			int sheetIndex = parsingBook.getExternalSheetIndex(null, sheetName); // create index according parsing book logic
			Ptg[] tokens = FormulaParser.parse(formula, parsingBook, FormulaType.CELL, sheetIndex); // current sheet index in parsing is always 0

			// move formula, limit to bound if dest. is out; if first and last both out on bound, it will be "#REF!"
			String regionBookName = region.getSheet().getBook().getBookName();
			String regionSheetName = region.getSheet().getSheetName();
			int regionSheetIndex;
			if(bookName.equals(regionBookName)) { // at same book
				regionSheetIndex = parsingBook.findExternalSheetIndex(regionSheetName); // find index, DON'T create
			} else { // different books
				regionSheetIndex = parsingBook.findExternalSheetIndex(regionBookName, regionSheetName); // find index, DON'T create
			}
			PtgShifter shifter = new PtgShifter(regionSheetIndex, region.getRow(), region.getLastRow(),
					rowOffset, region.getColumn(), region.getLastColumn(), columnOffset,
					parsingBook.getSpreadsheetVersion());
			boolean moved = shifter.adjustFormula(tokens, sheetIndex);

			// render formula, detect region and create result
			String renderedFormula = moved ? FormulaRenderer.toFormulaString(parsingBook, tokens) : formula;
			Ref singleRef = tokens.length == 1 ? toDenpendRef(context, parsingBook, tokens[0]) : null;
			expr = new FormulaExpressionImpl(renderedFormula, singleRef, false);

		} catch(FormulaParseException e) {
			logger.log(Level.INFO, e.getMessage());
			expr = new FormulaExpressionImpl(formula, null, true);
		}
		return expr;
	}

	@Override
	public FormulaExpression shrink(String formula, SheetRegion srcRegion, boolean horizontal, FormulaParseContext context) {
		NSheet sheet = context.getSheet();

		// shrinking is equals to move the neighbor region
		// calculate the neighbor and offset
		int rowOffset = 0, colOffset = 0;
		SheetRegion neighbor;
		if(horizontal) {
			// adjust on column
			colOffset = -srcRegion.getColumnCount();
			int col = srcRegion.getLastColumn() + 1;
			int lastCol = context.getBook().getMaxColumnIndex();
			// no change on row
			int row = srcRegion.getRow();
			int lastRow = srcRegion.getLastRow();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		} else { // vertical
			// adjust on row
			rowOffset = -srcRegion.getRowCount();
			int row = srcRegion.getLastRow() + 1;
			int lastRow = context.getBook().getMaxRowIndex();
			// no change on column
			int col = srcRegion.getColumn();
			int lastCol = srcRegion.getLastColumn();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		}

		// move it
		return move(formula, neighbor, rowOffset, colOffset, context);
	}

	@Override
	public FormulaExpression extend(String formula, SheetRegion srcRegion,
			boolean horizontal, FormulaParseContext context) {
		NSheet sheet = context.getSheet();

		// extending is equals to move selected region and neighbor region at the same time
		// calculate the target region (combined selected and neighbor) and offset
		int rowOffset = 0, colOffset = 0;
		SheetRegion neighbor;
		if(horizontal) {
			// adjust on column
			colOffset = srcRegion.getColumnCount();
			int col = srcRegion.getColumn();
			int lastCol = context.getBook().getMaxColumnIndex();
			// no change on row
			int row = srcRegion.getRow();
			int lastRow = srcRegion.getLastRow();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		} else { // vertical
			// adjust on row
			rowOffset = srcRegion.getRowCount();
			int row = srcRegion.getRow();
			int lastRow = context.getBook().getMaxRowIndex();
			// no change on column
			int col = srcRegion.getColumn();
			int lastCol = srcRegion.getLastColumn();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		}

		// move it
		return move(formula, neighbor, rowOffset, colOffset, context);
	}

	@Override
	public FormulaExpression shift(String formula, int rowOffset,
			int columnOffset, FormulaParseContext context) {
		
		FormulaExpression expr = null;
		try {
			// adapt and parse
			NBook book = context.getBook();
			ParsingBook parsingBook = new ParsingBook(book);
			String sheetName = context.getSheet().getSheetName();
			int sheetIndex = parsingBook.getExternalSheetIndex(null, sheetName); // create index according parsing book logic
			Ptg[] tokens = FormulaParser.parse(formula, parsingBook, FormulaType.CELL, sheetIndex); // current sheet index in parsing is always 0

			// simply shift every PTG and no need to consider sheet index
			if(rowOffset != 0 || columnOffset != 0) { // shift formula only if necessary
				for(int i = 0; i < tokens.length; ++i) {
					Ptg ptg = tokens[i];
					if(ptg instanceof RefPtgBase) {
						RefPtgBase rptg = (RefPtgBase)ptg;
						// calculate offset
						int r = rptg.getRow() + (rptg.isRowRelative() ? rowOffset : 0);
						int c = rptg.getColumn() + (rptg.isColRelative() ? columnOffset : 0);
						// if reference is out of bounds, convert it to #REF
						if(isValidRowIndex(book, r) && isValidColumnIndex(book, c)) {
							rptg.setRow(r);
							rptg.setColumn(c);
						} else {
							tokens[i] = PtgShifter.createDeletedRef(rptg);
						}
					} else if(ptg instanceof AreaPtgBase) {
						AreaPtgBase aptg = (AreaPtgBase)ptg;
						// shift
						int r0 = aptg.getFirstRow() + (aptg.isFirstRowRelative() ? rowOffset : 0);
						int r1 = aptg.getLastRow() + (aptg.isLastRowRelative() ? rowOffset : 0);
						int c0 = aptg.getFirstColumn() + (aptg.isFirstColRelative() ? columnOffset : 0);
						int c1 = aptg.getLastColumn() + (aptg.isLastColRelative() ? columnOffset : 0);
						// if reference is out of bounds, convert it to #REF
						if(isValidRowIndex(book, r0) && isValidRowIndex(book, r1)
								&& isValidColumnIndex(book, c0) && isValidColumnIndex(book, c1)) {
							aptg.setFirstRow(Math.min(r0, r1));
							aptg.setLastRow(Math.max(r0, r1));
							aptg.setFirstColumn(Math.min(c0, c1));
							aptg.setLastColumn(Math.max(c0, c1));
						} else {
							tokens[i] = PtgShifter.createDeletedRef(aptg);
						}
					}
				}
			}

			// render formula, detect region and create result
			String renderedFormula = FormulaRenderer.toFormulaString(parsingBook, tokens);
			Ref singleRef = tokens.length == 1 ? toDenpendRef(context, parsingBook, tokens[0]) : null;
			expr = new FormulaExpressionImpl(renderedFormula, singleRef, false);

		} catch(FormulaParseException e) {
			logger.log(Level.INFO, e.getMessage());
			expr = new FormulaExpressionImpl(formula, null, true);
		}
		return expr;
	}
	
	private boolean isValidRowIndex(NBook book, int rowIndex) {
		return 0 <= rowIndex && rowIndex <= book.getMaxRowIndex();
	}
	
	private boolean isValidColumnIndex(NBook book, int columnIndex) {
		return 0 <= columnIndex && columnIndex <= book.getMaxColumnIndex();
	}

	@Override
	public FormulaExpression transpose(String formula, int rowOrigin,
			int columnOrigin, FormulaParseContext context) {
		
		FormulaExpression expr = null;
		try {
			// adapt and parse
			NBook book = context.getBook();
			ParsingBook parsingBook = new ParsingBook(book);
			String sheetName = context.getSheet().getSheetName();
			int sheetIndex = parsingBook.getExternalSheetIndex(null, sheetName); // create index according parsing book logic
			Ptg[] tokens = FormulaParser.parse(formula, parsingBook, FormulaType.CELL, sheetIndex); // current sheet index in parsing is always 0

			// simply adjust every PTG and no need to consider sheet index
			for(int i = 0; i < tokens.length; ++i) {
				Ptg ptg = tokens[i];
				if(ptg instanceof RefPtgBase) {
					RefPtgBase rptg = (RefPtgBase)ptg;
					
					// process transpose only if both directions are relative
					if(rptg.isRowRelative() && rptg.isColRelative()) {
						
						// every direction:
						// 1. translate origin 2. swap row and column 3. translate origin back
						int r = (rptg.getColumn() - columnOrigin) + rowOrigin;
						int c = (rptg.getRow() - rowOrigin) + columnOrigin;
						
						// if reference is out of bounds, convert it to #REF
						if(isValidRowIndex(book, r) && isValidColumnIndex(book, c)) {
							rptg.setRow(r);
							rptg.setColumn(c);
						} else {
							tokens[i] = PtgShifter.createDeletedRef(rptg);
						}
					}
				} else if(ptg instanceof AreaPtgBase) {
					AreaPtgBase aptg = (AreaPtgBase)ptg;

					// need transpose process if ANY pair's both directions are relative
					// if so,
					// 1. this process skip any rest absolute setting
					// 2. swap absolute setting to another direction 
					if((aptg.isFirstRowRelative() && aptg.isFirstColRelative()) 
							|| (aptg.isLastRowRelative() && aptg.isLastColRelative())) {
						
						// every direction:
						// 1. translate origin 2. swap row and column 3. translate origin back
						int r0 = (aptg.getFirstColumn() - columnOrigin) + rowOrigin;
						int c0 = (aptg.getFirstRow() - rowOrigin) + columnOrigin;
						int r1 = (aptg.getLastColumn() - columnOrigin) + rowOrigin;
						int c1 = (aptg.getLastRow() - rowOrigin) + columnOrigin;
						
						// swap absolute setting
						boolean temp = aptg.isFirstRowRelative();
						aptg.setFirstRowRelative(aptg.isFirstColRelative());
						aptg.setFirstColRelative(temp);
						temp = aptg.isLastRowRelative();
						aptg.setLastRowRelative(aptg.isLastColRelative());
						aptg.setLastColRelative(temp);
						
						// if reference is out of bounds, convert it to #REF
						if(isValidRowIndex(book, r0) && isValidRowIndex(book, r1)
								&& isValidColumnIndex(book, c0) && isValidColumnIndex(book, c1)) {
							aptg.setFirstRow(Math.min(r0, r1));
							aptg.setLastRow(Math.max(r0, r1));
							aptg.setFirstColumn(Math.min(c0, c1));
							aptg.setLastColumn(Math.max(c0, c1));
						} else {
							tokens[i] = PtgShifter.createDeletedRef(aptg);
						}
					}
				}
			}

			// render formula, detect region and create result
			String renderedFormula = FormulaRenderer.toFormulaString(parsingBook, tokens);
			Ref singleRef = tokens.length == 1 ? toDenpendRef(context, parsingBook, tokens[0]) : null;
			expr = new FormulaExpressionImpl(renderedFormula, singleRef, false);

		} catch(FormulaParseException e) {
			logger.log(Level.INFO, e.getMessage());
			expr = new FormulaExpressionImpl(formula, null, true);
		}
		return expr;
	}

}
