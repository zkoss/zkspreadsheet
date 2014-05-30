/* ZPOIEngine.java

	Purpose:
		
	Description:
		
	History:
		Dec 10, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model.impl.sys.formula;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.zkoss.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.zkoss.poi.ss.formula.DependencyTracker;
import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationSheet;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.EvaluationWorkbook.ExternalSheet;
import org.zkoss.poi.ss.formula.ExternSheetReferenceToken;
import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.formula.FormulaRenderingWorkbook;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.IStabilityClassifier;
import org.zkoss.poi.ss.formula.PtgShifter;
import org.zkoss.poi.ss.formula.SheetNameFormatter;
import org.zkoss.poi.ss.formula.WorkbookDependentFormula;
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
import org.zkoss.poi.ss.formula.ptg.OperandPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.ptg.Ref3DPtg;
import org.zkoss.poi.ss.formula.ptg.RefPtg;
import org.zkoss.poi.ss.formula.ptg.RefPtgBase;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.util.LittleEndianOutput;
import org.zkoss.poi.xssf.model.IndexedUDFFinder;
import org.zkoss.util.logging.Log;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelContext;
import org.zkoss.xel.util.SimpleXelContext;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.impl.NameRefImpl;
import org.zkoss.zss.model.impl.NonSerializableHolder;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.impl.sys.DependencyTableAdv;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.dependency.Ref.RefType;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.FormulaClearContext;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
import org.zkoss.zss.model.sys.formula.FunctionResolver;
import org.zkoss.zss.model.sys.formula.FunctionResolverFactory;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;

/**
 * A formula engine implemented by ZPOI
 * @author Pao
 */
public class FormulaEngineImpl implements FormulaEngine {

	public final static String KEY_EVALUATORS = "$ZSS_EVALUATORS$";

	private static final Log _logger = Log.lookup(FormulaEngineImpl.class.getName());

	private Map<EvaluationWorkbook, XelContext> _xelContexts = new HashMap<EvaluationWorkbook, XelContext>();
	
	// for POI formula evaluator
	protected final static IStabilityClassifier noCacheClassifier = new IStabilityClassifier() {
		public boolean isCellFinal(int sheetIndex, int rowIndex, int columnIndex) {
			return true;
		}
	};
	
	private static Pattern _areaPattern = Pattern.compile("\\([.[^\\(\\)]]*\\)");//match (A1,B1,C1)
	private static Pattern _searchPattern = Pattern.compile("\\s*((?:(?:'[^!\\(]+'!)|(?:[^'!,\\(]+!))?(?:[$\\w]+:)?[$\\w]+)"); // for search area reference 
	
	private static boolean isMultipleAreaFormula(String formula){
		return _areaPattern.matcher(formula).matches();
		
	}
	
	private String[] unwrapeAreaFormula(String formula){
		List<String> areaStrings = new ArrayList<String>();
		Matcher m = _searchPattern.matcher(formula);
		while(m.find()) {
			areaStrings.add(m.group(1));
		}
		return areaStrings.toArray(new String[0]);
	}


	private FormulaExpression parseMultipleAreaFormula(String formula, FormulaParseContext context) {
		if(!isMultipleAreaFormula(formula)){
			return null;
		}
		FormulaExpression[] result = parse0(unwrapeAreaFormula(formula),context);
		
		List<Ref> areaRefs = new LinkedList<Ref>(); 
		for(FormulaExpression expr:result){
			if(expr.hasError()){
				return new FormulaExpressionImpl(formula, null, true,expr.getErrorMessage());
			}
			for(Ref ref:expr.getAreaRefs()){
				areaRefs.add(ref);
			}
		}
		return new FormulaExpressionImpl(formula, areaRefs.toArray(new Ref[areaRefs.size()]));
	}
	@Override
	public FormulaExpression parse(String formula, FormulaParseContext context) {
		formula = formula.trim();
		FormulaExpression expr = parseMultipleAreaFormula(formula,context);
		if(expr!=null){
			return expr;
		}
		return parse0(new String[]{formula},context)[0];
	}

	private FormulaExpression[] parse0(String[] formulas, FormulaParseContext context) {
		LinkedList<FormulaExpression> result = new LinkedList<FormulaExpression>();
		Ref dependant = context.getDependent();
		LinkedList<Ref> precednets = dependant!=null?new LinkedList<Ref>():null;
		try {
			// adapt and parse
			SBook book = context.getBook();
			ParsingBook parsingBook = new ParsingBook(book);
			int sheetIndex = parsingBook.getExternalSheetIndex(null, context.getSheet().getSheetName());
			
			AbstractBookSeriesAdv series = (AbstractBookSeriesAdv)book.getBookSeries();
			DependencyTableAdv table = (DependencyTableAdv)series.getDependencyTable();
			
			boolean error = false;
			
			for(String formula:formulas){
				try{
					Ptg[] tokens = parse(formula, parsingBook, sheetIndex, context); // current sheet index in parsing is always 0
					if(dependant!=null){
						for(Ptg ptg : tokens) {
							Ref precedent = toDenpendRef(context, parsingBook, ptg);
							if(precedent != null) {
								precednets.add(precedent);
							}
						}
					}
					
					// render formula, detect region and create result
					String renderedFormula = renderFormula(parsingBook, formula, tokens, false);
					Ref singleRef = tokens.length == 1 ? toDenpendRef(context, parsingBook, tokens[0]) : null;
					Ref[] refs = singleRef==null ? null :
						(singleRef.getType() == RefType.AREA || singleRef.getType() == RefType.CELL ?new Ref[]{singleRef}:null);
					result.add(new FormulaExpressionImpl(renderedFormula, refs));
					
					
				}catch(FormulaParseException e) {
					_logger.info(e.getMessage() + " when parsing " + formula);
					result.add(new FormulaExpressionImpl(formula, null, true,e.getMessage()));
					error = true;
				} catch(Exception e) {
					_logger.error(e.getMessage() + " when parsing " + formula, e);
					result.add(new FormulaExpressionImpl(formula, null, true,e.getMessage()));
					error = true;
				}
			}
			
			// dependency tracking if no error and necessary
			if(!error && dependant != null) {
				for(Ref precedent:precednets){
					table.add(dependant, precedent);
				}
			}
		} catch(Exception e) {
			_logger.error(e.getMessage() + " when parsing " + Arrays.asList(formulas), e);
			result.clear();
			result.add(new FormulaExpressionImpl(Arrays.asList(formulas).toString(), null, true,e.getMessage()));
		}
		return result.toArray(new FormulaExpression[result.size()]);
	}	
	
	protected Ptg[] parse(String formula, FormulaParsingWorkbook book, int sheetIndex, FormulaParseContext context) {
		return FormulaParser.parse(formula, book, FormulaType.CELL, sheetIndex);
	}
	
	protected String renderFormula(ParsingBook parsingBook, String formula, Ptg[] tokens, boolean always) {
		if (always) return FormulaRenderer.toFormulaString(parsingBook, tokens);
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
			SSheet sheet = ctx.getSheet();

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
			_logger.error(e.getMessage(), e);
		}
		return null;
	}

	@Override
	@SuppressWarnings("unchecked")
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context) {

		// by pass if expression is invalid format
		if(expr.hasError()) {
			return new EvaluationResultImpl(ResultType.ERROR,  ErrorValue.valueOf(ErrorValue.INVALID_FORMULA));
		}
		Ref dependant = context.getDependent();
		EvaluationResult result = null;
		try {

			// get evaluation context from book series
			SBook book = context.getBook();
			AbstractBookSeriesAdv bookSeries = (AbstractBookSeriesAdv)book.getBookSeries();
			DependencyTableAdv table = (DependencyTableAdv)bookSeries.getDependencyTable();	
			NonSerializableHolder<Map<String, EvalContext>> holder = (NonSerializableHolder<Map<String, EvalContext>>)bookSeries
					.getAttribute(KEY_EVALUATORS);
			Map<String, EvalContext> evalCtxMap = holder == null?null:holder.getObject();

			// get evaluation context or create new one if not existed
			if(evalCtxMap == null) {
				evalCtxMap = new LinkedHashMap<String, FormulaEngineImpl.EvalContext>();
				List<String> bookNames = new ArrayList<String>();
				List<WorkbookEvaluator> evaluators = new ArrayList<WorkbookEvaluator>();
				for(SBook nb : bookSeries.getBooks()) {
					String bookName = nb.getBookName();
					EvalBook evalBook = new EvalBook(nb);
					WorkbookEvaluator we = new WorkbookEvaluator(evalBook, noCacheClassifier, null);
					bookNames.add(bookName);
					evaluators.add(we);
					evalCtxMap.put(bookName, new EvalContext(evalBook, we));

					// aggregate built-in functions and user defined functions
					FunctionResolver resolver = FunctionResolverFactory.createFunctionResolver();
					UDFFinder zkUDFF = resolver.getUDFFinder(); // ZK user defined function finder
					if(zkUDFF != null) {
						IndexedUDFFinder bookUDFF = (IndexedUDFFinder)evalBook.getUDFFinder(); // book contained built-in function finder
						bookUDFF.insert(0, zkUDFF);
					}
				}
				CollaboratingWorkbooksEnvironment.setup(bookNames.toArray(new String[0]),
						evaluators.toArray(new WorkbookEvaluator[0]));
				bookSeries.setAttribute(KEY_EVALUATORS, new NonSerializableHolder<Map<String, EvalContext>>(evalCtxMap));
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
			if(dependant!=null){
				table.setEvaluated(dependant);
			}
		} catch(NotImplementedException e) {
			_logger.info(e.getMessage() + " when eval " + expr.getFormulaString());
			result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_NAME, e.getMessage()));
		} catch(EvaluationException e) { 
			_logger.warning(e.getMessage() + " when eval " + expr.getFormulaString());
			ErrorEval error = e.getErrorEval();
			result = new EvaluationResultImpl(ResultType.ERROR, error==null?new ErrorValue(ErrorValue.INVALID_FORMULA, e.getMessage()):error);
		} catch(FormulaParseException e) {
			// we skip evaluation if formula has parsing error
			// so if still occurring formula parsing exception, it should be a bug 
			_logger.error(e.getMessage() + " when eval " + expr.getFormulaString());
			result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_FORMULA, e.getMessage()));
		} catch(Exception e) {
			_logger.error(e.getMessage() + " when eval " + expr.getFormulaString(), e);
			result = new EvaluationResultImpl(ResultType.ERROR, new ErrorValue(ErrorValue.INVALID_FORMULA, e.getMessage()));
		}
		return result;
	}

	protected EvaluationResult evaluateFormula(FormulaExpression expr, FormulaEvaluationContext context, EvalBook evalBook, WorkbookEvaluator evaluator) throws FormulaParseException, Exception {

		// do evaluate
		SBook book = context.getBook();
		int currentSheetIndex = book.getSheetIndex(context.getSheet());
		SCell cell = context.getCell();
		ValueEval value = null;
		boolean multipleArea = isMultipleAreaFormula(expr.getFormulaString());
		if(cell == null || cell.isNull()) {
			// evaluation formula directly
			if(multipleArea){
				String[] formulas = unwrapeAreaFormula(expr.getFormulaString());
				List<ValueEval> evals = new ArrayList<ValueEval>(formulas.length);
				for(String f:formulas){
					value = evaluator.evaluate(currentSheetIndex, f, true);
					evals.add(value);
				}
				value = new ValuesEval(evals.toArray(new ValueEval[evals.size()]));
			}else{
				value = evaluator.evaluate(currentSheetIndex, expr.getFormulaString(), true);
			}
		} else {
			if(multipleArea){//is multipleArea formula in cell, should return #VALUE!
				return new EvaluationResultImpl(ResultType.ERROR, ErrorValue.valueOf(ErrorValue.INVALID_VALUE));
			}
			EvaluationCell evalCell = evalBook.getSheet(currentSheetIndex).getCell(cell.getRowIndex(),
					cell.getColumnIndex());
			value = evaluator.evaluate(evalCell);
		}

		// convert to result
		if(value instanceof ErrorEval) {
			int code = ((ErrorEval)value).getErrorCode();
			return new EvaluationResultImpl(ResultType.ERROR, ErrorValue.valueOf((byte)code));
		} else {
			try{
				return new EvaluationResultImpl(ResultType.SUCCESS, getResolvedValue(value));
			}catch(EvaluationException x){
				//error when resolve value.
				if(x.getErrorEval()!=null){//ZSS-591 Get console exception after delete sheet
					return new EvaluationResultImpl(ResultType.ERROR, ErrorValue.valueOf((byte)x.getErrorEval().getErrorCode()));
				}else{
					throw x;
				}
			}
		}
	}

	protected Object getResolvedValue(ValueEval value) throws EvaluationException {
		if(value instanceof StringEval) {
			return ((StringEval)value).getStringValue();
		} else if(value instanceof NumberEval) {
			return ((NumberEval)value).getNumberValue();
		} else if(value instanceof BlankEval) {
			return "";
		} else if(value instanceof BoolEval) {
			return ((BoolEval)value).getBooleanValue();
		} else if(value instanceof ValuesEval) {
			ValueEval[] values = ((ValuesEval)value).getValueEvals();
			Object[] array = new Object[values.length];
			for(int i = 0; i < values.length; ++i) {
				array[i] = getResolvedValue(values[i]);
			}
			return array;
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
			return list;
		} else if(value instanceof RefEval) {
			ValueEval ve = ((RefEval)value).getInnerValueEval();
			Object v = getResolvedValue(ve);
			return v;
		} else if(value instanceof ErrorEval) {
			throw new EvaluationException((ErrorEval)value);
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
		XelContext xelContext = _xelContexts.get(evalBook);
		if(xelContext == null) {

			// create function resolver
			FunctionResolver resolver = FunctionResolverFactory.createFunctionResolver();

			// apply POI dependency tracker for defined name resolving
			DependencyTracker tracker = resolver.getDependencyTracker();
			if(tracker != null) {
				evaluator.setDependencyTracker(tracker);
			}
			
			// collect all function mappers 
			JoinFunctionMapper functionMapper = new JoinFunctionMapper(null);
			FunctionMapper extraFunctionMapper = context.getFunctionMapper();
			if(extraFunctionMapper != null) {
				functionMapper.addFunctionMapper(extraFunctionMapper); // must before ZSS function mapper
			}
			FunctionMapper zssFuncMapper = resolver.getFunctionMapper();
			if(zssFuncMapper != null) {
				functionMapper.addFunctionMapper(zssFuncMapper);
			}
			
			// collect all variable resolvers
			JoinVariableResolver variableResolver = new JoinVariableResolver();
			VariableResolver extraVariableResolver = context.getVariableResolver();
			if(extraVariableResolver != null) {
				variableResolver.addVariableResolver(extraVariableResolver);
			}
			
			// create XEL context
			xelContext = new SimpleXelContext(variableResolver, functionMapper);
			xelContext.setAttribute("zkoss.zss.CellType", Object.class);
			_xelContexts.put(evalBook, xelContext);
		}
		return xelContext;
	}

	@Override
	@SuppressWarnings("unchecked")
	public void clearCache(FormulaClearContext context) {
		try {
			SBook book = context.getBook();
			SSheet sheet = context.getSheet();
			SCell cell = context.getCell();

			// take evaluators from book series
			AbstractBookSeriesAdv bookSeries = (AbstractBookSeriesAdv)book.getBookSeries();
			NonSerializableHolder<Map<String, EvalContext>> holder = (NonSerializableHolder<Map<String, EvalContext>>)bookSeries
					.getAttribute(KEY_EVALUATORS);
			Map<String, EvalContext> map = holder == null?null:holder.getObject();

			// do nothing if not existed
			if(map == null) {
				return;
			}

			// clean cache and target is a cell
			if(cell != null && !cell.isNull()) {

				// do nothing if not existed
				EvalContext ctx = map.get(book.getBookName());
				if(ctx == null) {
					_logger.warning("clear a non-existed book? >> " + book.getBookName());
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
			_logger.error(e.getMessage(), e);
		}
	}

	protected static class FormulaExpressionImpl implements FormulaExpression, Serializable {
		private static final long serialVersionUID = -8532826169759927711L;
		private String formula;
		private Ref[] refs;
		private boolean error;
		private String errorMessage;

		/**
		 * @param ref resolved reference if formula has only one parsed token
		 */
		public FormulaExpressionImpl(String formula, Ref[] refs) {
			this(formula,refs,false,null);
		}
		public FormulaExpressionImpl(String formula, Ref[] refs, boolean error, String errorMessage) {
			this.formula = formula;
			if(refs!=null){
				for(Ref ref:refs){
					if( ref.getType() == RefType.AREA || ref.getType() == RefType.CELL){
						continue;
					}
					this.error = true;
					this.errorMessage = errorMessage==null?"wrong area reference":errorMessage;
					return;
				}
			}
			this.refs = refs;
			this.error = error;
			this.errorMessage = errorMessage;
			
		}

		@Override
		public boolean hasError() {
			return error;
		}
		
		@Override
		public String getErrorMessage(){
			return errorMessage;
		}

		@Override
		public String getFormulaString() {
			return formula;
		}

		@Override
		public boolean isAreaRefs() {
			return refs != null && refs.length>0;
		}
		
		@Override
		public Ref[] getAreaRefs() {
			return refs;
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
	
	private FormulaExpression adjustMultipleArea(String formula, FormulaParseContext context, FormulaAdjuster adjuster) {
		if(!isMultipleAreaFormula(formula)){
			return null;
		}
		StringBuilder sb = new StringBuilder("(");
		//handle multiple area
		String[] fs = unwrapeAreaFormula(formula); 
		List<Ref> areaRefs = new LinkedList<Ref>();
		for(int i=0;i<fs.length;i++){
			if(i>0){
				sb.append(",");
			}
			FormulaExpression expr = adjust(fs[i], context, adjuster);
			if(expr.hasError()){
				return new FormulaExpressionImpl(formula, null,true,expr.getErrorMessage());
			}
			sb.append(expr.getFormulaString());
			if(expr.isAreaRefs()){
				for(Ref ref:expr.getAreaRefs()){
					areaRefs.add(ref);
				}
			}
		}
		sb.append(")");
		return new FormulaExpressionImpl(sb.toString(),areaRefs.size()==0?null:areaRefs.toArray(new Ref[areaRefs.size()]));
	}
	
	/**
	 * adjust formula through specific adjuster
	 */
	private FormulaExpression adjust(String formula, FormulaParseContext context, FormulaAdjuster adjuster) {
		FormulaExpression expr = null;
		try {
			// adapt and parse
			ParsingBook parsingBook = new ParsingBook(context.getBook());
			String sheetName = context.getSheet().getSheetName();
			int sheetIndex = parsingBook.getExternalSheetIndex(null, sheetName); // create index according parsing book logic
			Ptg[] tokens = FormulaParser.parse(formula, parsingBook, FormulaType.CELL, sheetIndex); // current sheet index in parsing is always 0

			// adjust formula
			boolean modified = adjuster.process(formula, sheetIndex, tokens, parsingBook, context);
			
			// render formula, detect region and create result
			String renderedFormula = modified ? 
					renderFormula(parsingBook, formula, tokens, true) : formula;
			Ref singleRef = tokens.length == 1 ? toDenpendRef(context, parsingBook, tokens[0]) : null;
			Ref[] refs = singleRef==null ? null :
				(singleRef.getType() == RefType.AREA || singleRef.getType() == RefType.CELL ?new Ref[]{singleRef}:null);
			expr = new FormulaExpressionImpl(renderedFormula, refs);

		} catch(FormulaParseException e) {
			_logger.info(e.getMessage());
			expr = new FormulaExpressionImpl(formula, null, true, e.getMessage());
		}
		return expr;
	}
	
	private static interface FormulaAdjuster {
		/**
		 * @return true if formula modified, denote this formula needs re-render
		 */
		public boolean process(String formula, int sheetIndex, Ptg[] tokens, ParsingBook parsingBook, FormulaParseContext context);
	}
	
	@Override
	public FormulaExpression move(String formula, final SheetRegion region, final int rowOffset, final int columnOffset, FormulaParseContext context) {
		formula = formula.trim();
		FormulaAdjuster shiftAdjuster = new FormulaAdjuster() {
			@Override
			public boolean process(String formula, int sheetIndex, Ptg[] tokens, ParsingBook parsingBook, FormulaParseContext context) {
				// move formula, limit to bound if dest. is out; if first and last both out on bound, it will be "#REF!"
				String bookName = context.getBook().getBookName();
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
				return shifter.adjustFormula(tokens, sheetIndex);
			}
		};
		FormulaExpression result = adjustMultipleArea(formula, context, shiftAdjuster);
		if(result!=null){
			return result;
		}
		return adjust(formula, context, shiftAdjuster);
	}

	@Override
	public FormulaExpression shrink(String formula, SheetRegion srcRegion, boolean horizontal, FormulaParseContext context) {
		SSheet sheet = srcRegion.getSheet();

		// shrinking is equals to move the neighbor region
		// calculate the neighbor and offset
		int rowOffset = 0, colOffset = 0;
		SheetRegion neighbor;
		if(horizontal) {
			// adjust on column
			colOffset = -srcRegion.getColumnCount();
			int col = srcRegion.getLastColumn() + 1;
			int lastCol = sheet.getBook().getMaxColumnIndex();
			// no change on row
			int row = srcRegion.getRow();
			int lastRow = srcRegion.getLastRow();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		} else { // vertical
			// adjust on row
			rowOffset = -srcRegion.getRowCount();
			int row = srcRegion.getLastRow() + 1;
			int lastRow = sheet.getBook().getMaxRowIndex();
			// no change on column
			int col = srcRegion.getColumn();
			int lastCol = srcRegion.getLastColumn();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		}

		// move it
		return move(formula, neighbor, rowOffset, colOffset, context);
	}

	@Override
	public FormulaExpression extend(String formula, SheetRegion srcRegion, boolean horizontal, FormulaParseContext context) {
		SSheet sheet = srcRegion.getSheet();

		// extending is equals to move selected region and neighbor region at the same time
		// calculate the target region (combined selected and neighbor) and offset
		int rowOffset = 0, colOffset = 0;
		SheetRegion neighbor;
		if(horizontal) {
			// adjust on column
			colOffset = srcRegion.getColumnCount();
			int col = srcRegion.getColumn();
			int lastCol = sheet.getBook().getMaxColumnIndex();
			// no change on row
			int row = srcRegion.getRow();
			int lastRow = srcRegion.getLastRow();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		} else { // vertical
			// adjust on row
			rowOffset = srcRegion.getRowCount();
			int row = srcRegion.getRow();
			int lastRow = sheet.getBook().getMaxRowIndex();
			// no change on column
			int col = srcRegion.getColumn();
			int lastCol = srcRegion.getLastColumn();
			neighbor = new SheetRegion(sheet, row, col, lastRow, lastCol);
		}

		// move it
		return move(formula, neighbor, rowOffset, colOffset, context);
	}

	@Override
	public FormulaExpression shift(String formula, final int rowOffset, final int columnOffset, FormulaParseContext context) {
		formula = formula.trim();
		FormulaAdjuster shiftAdjuster = new FormulaAdjuster() {
			@Override
			public boolean process(String formula, int sheetIndex, Ptg[] tokens, ParsingBook parsingBook, FormulaParseContext context) {
				
				// shift formula only if necessary
				if(rowOffset != 0 || columnOffset != 0) {
					
					// simply shift every PTG and no need to consider sheet index
					SBook book = context.getBook();
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
					return true;
				} else {
					return false;
				}
			}
		};
		FormulaExpression result = adjustMultipleArea(formula, context, shiftAdjuster);
		if(result!=null){
			return result;
		}
		return adjust(formula, context, shiftAdjuster);
	}
	
	private boolean isValidRowIndex(SBook book, int rowIndex) {
		return 0 <= rowIndex && rowIndex <= book.getMaxRowIndex();
	}
	
	private boolean isValidColumnIndex(SBook book, int columnIndex) {
		return 0 <= columnIndex && columnIndex <= book.getMaxColumnIndex();
	}

	@Override
	public FormulaExpression transpose(String formula, final int rowOrigin, final int columnOrigin, FormulaParseContext context) {
		formula = formula.trim();
		FormulaAdjuster shiftAdjuster = new FormulaAdjuster() {
			@Override
			public boolean process(String formula, int sheetIndex, Ptg[] tokens, ParsingBook parsingBook, FormulaParseContext context) {

				// simply adjust every PTG and no need to consider sheet index
				SBook book = context.getBook();
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
				return true;
			}
		};
		
		FormulaExpression result = adjustMultipleArea(formula, context, shiftAdjuster);
		if(result!=null){
			return result;
		}
		return adjust(formula, context, shiftAdjuster);
	}

	@Override
	public FormulaExpression renameSheet(String formula, final SBook targetBook, final String oldSheetName, final String newSheetName, FormulaParseContext context) {
		formula = formula.trim();
		FormulaAdjuster shiftAdjuster = new FormulaAdjuster() {
			@Override
			public boolean process(String formula, int formulaSheetIndex, Ptg[] tokens, ParsingBook parsingBook, FormulaParseContext context) {
				if(newSheetName != null) {
					// parsed tokens has only external sheet index, not real sheet name
					// the sheet names are kept in parsing book, so we just rename sheets in parsing book
					// finally use such parsing book to re-render formula will get a renamed formula
					parsingBook.renameSheet(targetBook.getBookName(), oldSheetName, newSheetName);
					
				} else { // if new sheet name is null, it indicates deleted sheet
					
					// compare every token and replace it by deleted reference if necessary
					String bookName = targetBook == context.getBook() ? null : targetBook.getBookName(); 
					for(int i = 0; i < tokens.length; ++i) {
						Ptg ptg = tokens[i];
						if(ptg instanceof ExternSheetReferenceToken) { // must be Ref3DPtg or Area3DPtg 
							ExternSheetReferenceToken t = (ExternSheetReferenceToken)ptg;
							ExternalSheet es = parsingBook.getAnyExternalSheet(t.getExternSheetIndex());
							if((bookName == null && es.getWorkbookName() == null)
									|| (bookName != null && bookName.equals(es.getWorkbookName()))) {
								// replace token if any sheet name is matched 
								if(oldSheetName.equals(es.getSheetName()) || oldSheetName.equals(es.getLastSheetName())) {
									tokens[i] = new DeletedSheet3DPtg(bookName, ptg);
								}
							}
						}
					}
				}
				return true;
			}
		};
		FormulaExpression result = adjustMultipleArea(formula, context, shiftAdjuster);
		if(result!=null){
			return result;
		}
		return adjust(formula, context, shiftAdjuster);
	}
	
	private static class DeletedSheet3DPtg extends OperandPtg implements WorkbookDependentFormula {
		private OperandPtg ptg;
		private String bookName;

		public DeletedSheet3DPtg(String bookName, Ptg ptg) {
			if(ptg instanceof Ref3DPtg || ptg instanceof Area3DPtg) {
				this.ptg = (OperandPtg)ptg;
				this.bookName = bookName;
			} else {
				throw new IllegalArgumentException("must be Ref3dPtg or Area3DPtg");
			}
		}

		@Override
		public int getSize() {
			return ptg.getSize();
		}

		@Override
		public void write(LittleEndianOutput out) {
			// do nothing
		}

		@Override
		public String toFormulaString() {
			return ptg.toFormulaString();
		}

		@Override
		public byte getDefaultOperandClass() {
			return ptg.getDefaultOperandClass();
		}

		@Override
		public String toFormulaString(FormulaRenderingWorkbook book) {
			StringBuffer sb = new StringBuffer();
			if(bookName != null) {
				SheetNameFormatter.appendFormat(sb, bookName, "#REF");
			} else {
				// because of the POI parser's limitation, it can't parse #REF!A1
				// we use unusual '#REF' sheet name to represent a deleted sheet
				SheetNameFormatter.appendFormat(sb, "#REF");
				// sb.append("#REF"); // don't use SheetNameFormatter, it will add quote because of #
			}
			return sb.append('!').append(((ExternSheetReferenceToken)ptg).format2DRefAsString()).toString();
		}

		@Override
		public String toInternalFormulaString(FormulaRenderingWorkbook book) {
			throw new UnsupportedOperationException("unsupport internal formula string");
		}
	};
}
