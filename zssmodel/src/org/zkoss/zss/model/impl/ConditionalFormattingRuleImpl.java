/* ConditonalFormattingRuleImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 10:16:50 AM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.text.Format;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.zkoss.lang.Strings;
import org.zkoss.poi.ss.formula.ptg.AreaPtg;
import org.zkoss.poi.ss.formula.ptg.BoolPtg;
import org.zkoss.poi.ss.formula.ptg.ErrPtg;
import org.zkoss.poi.ss.formula.ptg.IntPtg;
import org.zkoss.poi.ss.formula.ptg.NumberPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.ptg.RefPtg;
import org.zkoss.poi.ss.formula.ptg.StringPtg;
import org.zkoss.poi.ss.usermodel.DataFormatter;
import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.InvalidFormulaException;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SColorScale;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SDataBar;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SIconSet;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.sys.DependencyTableAdv;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.model.sys.input.InputEngine;
import org.zkoss.zss.model.sys.input.InputParseContext;
import org.zkoss.zss.model.sys.input.InputResult;
import org.zkoss.zss.model.util.Validations;
import org.zkoss.zss.range.impl.CellMatch2;
import org.zkoss.zss.range.impl.ContainsBlank;
import org.zkoss.zss.range.impl.ContainsError;
import org.zkoss.zss.range.impl.DatesMatch2;
import org.zkoss.zss.range.impl.DuplicateCell;
import org.zkoss.zss.range.impl.ExpressionMatch;
import org.zkoss.zss.range.impl.Matchable;
import org.zkoss.zss.range.impl.NumMatch;
import org.zkoss.zss.range.impl.TrueMatch;

/**
 * @author henri
 * @since 3.8.2
 * 
 */
//ZSS-1138
public class ConditionalFormattingRuleImpl implements SConditionalFormattingRule, Serializable {
	private static final long serialVersionUID = 5733467761359067350L;
	
	private RuleType type;
	private RuleOperator operator;
	private Integer priority;
	private SExtraStyle style;
	private boolean stopIfTrue;
	private RuleTimePeriod timePeriod;
	private Long rank;
	private boolean percent;
	private boolean bottom;
	private SColorScale colorScale;
	private SDataBar dataBar;
	private SIconSet iconSet;
	private String text;
	private boolean notAboveAverage;
	private boolean equalAverage;
	private Integer standardDeviation; //1 ~ 3

	private FormulaExpression _formula1Expr; //ZSS-1142
	private FormulaExpression _formula2Expr; //ZSS-1142
	private FormulaExpression _formula3Expr; //ZSS-1142

	private Matchable<SCell> _matchable; //ZSS-1142
	private CFValueObjectHelper _voHelper; //ZSS-1142
	
	// cached information for colorScale/dataBar/iconSet
	private Double _min;
	private Double _middle;
	private Double _max;
	private IconSetInfo[] _iconSetInfos;
	private Map<Double, SFill> _colorScaleCache;
	private RuleInfo _ruleInfo1;
	private RuleInfo _ruleInfo2;
	
	private SConditionalFormatting _cfi;
	
	public ConditionalFormattingRuleImpl(SConditionalFormatting cfi) {
		this._cfi = cfi;
	}
	
	//ZSS-1142
	@Override
	public SConditionalFormatting getFormatting() {
		return this._cfi;
	}
	
	//ZSS-1142
	@Override
	public SSheet getSheet() {
		return _cfi.getSheet();
	}
	
	//ZSS-1251
	@Override
	public void clearFormulaResultCache() {
		_matchable = null;
		_min = _max = _middle = null;
		_iconSetInfos = null;
		_voHelper = null;
		_colorScaleCache = null;
		_ruleInfo1 = _ruleInfo2 = null;
	}
	
	@Override
	public RuleType getType() {
		return type;
	}

	public void setType(RuleType type) {
		this.type = type;
	}
	
	@Override
	public RuleOperator getOperator() {
		return operator;
	}
	
	public void setOperator(RuleOperator operator) {
		this.operator = operator; 
	}

	@Override
	public Integer getPriority() {
		return priority;
	}
	
	public void setPriority(int priority) {
		this.priority = priority;
	}

	@Override
	public SExtraStyle getExtraStyle() {
		return style;
	}
	
	public void setExtraStyle(SExtraStyle style) {
		this.style = style;
	}

	@Override
	public boolean isStopIfTrue() {
		return stopIfTrue;
	}
	
	public void setStopIfTrue(boolean b) {
		stopIfTrue = b;
	}

	@Override
	public List<String> getFormulas() {
		List<String> fms = new ArrayList<String>();
		if (getFormula1() != null) {
			fms.add(getFormula1());
			if (getFormula2() != null) {
				fms.add(getFormula2());
				if (getFormula3() != null) {
					fms.add(getFormula3());
				}
			}
		}
		return fms;
	}
	
	@Override
	public RuleTimePeriod getTimePeriod() {
		return timePeriod;
	}
	
	public void setTimePeriod(RuleTimePeriod timePeriod) {
		this.timePeriod = timePeriod;
	}

	@Override
	public Long getRank() {
		return rank;
	}
	
	public void setRank(long rank) {
		this.rank = rank;
	}

	@Override
	public boolean isPercent() {
		return percent;
	}
	
	public void setPercent(boolean b) {
		percent = b;
	}

	@Override
	public boolean isBottom() {
		return bottom;
	}

	public void setBottom(boolean b) {
		bottom = b;
	}
	
	@Override
	public SColorScale getColorScale() {
		return colorScale;
	}
	
	public void setColorScale(SColorScale colorScale) {
		this.colorScale = colorScale;
	}

	@Override
	public SDataBar getDataBar() {
		return dataBar;
	}

	public void setDataBar(SDataBar dataBar) {
		this.dataBar = dataBar;
	}
	
	@Override
	public SIconSet getIconSet() {
		return iconSet;
	}

	public void setIconSet(SIconSet iconSet) {
		this.iconSet = iconSet;
	}
	
	@Override
	public String getText() {
		return text;
	}
	
	public void setText(String text) {
		this.text = text;
	}

	@Override
	public boolean isAboveAverage() {
		return !notAboveAverage;
	}
	
	public void setAboveAverage(boolean b) {
		notAboveAverage = !b;
	}

	@Override
	public boolean isEqualAverage() {
		return equalAverage;
	}
	
	public void setEqualAverage(boolean b) {
		equalAverage = b;
	}

	@Override
	public Integer getStandardDeviation() {
		return standardDeviation;
	}
	
	public void setStandardDeviation(int s) {
		standardDeviation = s;
	}

	//ZSS-1142
	// check if the provided cell and value match this conditional rule.
	public boolean match(SCell cell) {
		if (_matchable != null) {
			return _matchable.match(cell);
		}
		
		// Here we borrow the code from AutoFilter
		RuleOperator op1 = null;
		RuleOperator op2 = null;
		RuleOperator op = getOperator();
		boolean isAnd = true;
		List<SCFValueObject> vos = null;

		switch(this.type) {
		case BEGINS_WITH:
		case ENDS_WITH:
		case CONTAINS_TEXT:
		case NOT_CONTAINS_TEXT:
			op1 = op;
			op2 = null;
			if (this.text != null) {
				_matchable = new CellMatch2(op1, this.text);
			} else {
				final CellRegion region = getRegions().iterator().next();
				final int row0 = region.row;
				final int col0 = region.column;
				_ruleInfo1 = new RuleInfo(getSheet(), this, _formula2Expr, row0, col0);
				_matchable = new CellMatch2(op1, _ruleInfo1, op2, null, isAnd);
			}
			break;
		case CELL_IS:
		{
			switch (op) {
			case BETWEEN:
				op1 = RuleOperator.GREATER_THAN_OR_EQUAL;
				op2 = RuleOperator.LESS_THAN_OR_EQUAL;
				break;
			case NOT_BETWEEN:
				op1 = RuleOperator.LESS_THAN;
				op2 = RuleOperator.GREATER_THAN;
				isAnd = false;
				break;
			default:
				op1 = op;
				op2 = null;
				break;
			}
			final CellRegion region = getRegions().iterator().next();
			final int row0 = region.row;
			final int col0 = region.column;
			_ruleInfo1 = new RuleInfo(getSheet(), this, _formula1Expr, row0, col0);
			_ruleInfo2 = op2 == null ? null : new RuleInfo(getSheet(), this, _formula2Expr, row0, col0);
			_matchable = new CellMatch2(op1, _ruleInfo1, op2, _ruleInfo2, isAnd);
			break;
		}	
		case CONTAINS_BLANKS:
			_matchable = new ContainsBlank(false);
			break;
			
		case NOT_CONTAINS_BLANKS:
			_matchable = new ContainsBlank(true);
			break;
			
		case CONTAINS_ERRORS:
			_matchable = new ContainsError(false);
			break;
			
		case NOT_CONTAINS_ERRORS:
			_matchable = new ContainsError(true);
			break;
			
		case ABOVE_AVERAGE:
		{
			final double[] avg_stdv = _calcAverageStdv(); // consider stdDev
			final double avg = avg_stdv[0];
			final double stdv = avg_stdv[1] * (standardDeviation == null ? 0 : standardDeviation.intValue());
			final RuleOperator opx = notAboveAverage ?
					(equalAverage ?
						RuleOperator.LESS_THAN_OR_EQUAL:
						RuleOperator.LESS_THAN):
					(equalAverage ?
						RuleOperator.GREATER_THAN_OR_EQUAL:
						RuleOperator.GREATER_THAN);
			final double val = notAboveAverage ? // below
					avg - stdv : avg + stdv;
			final CellRegion region = getRegions().iterator().next();
			final int row0 = region.row;
			final int col0 = region.column;
			final FormulaExpression expr = parseFormula(""+val);
			_ruleInfo1 = new RuleInfo(getSheet(), this, expr, row0, col0);
			_matchable = new CellMatch2(opx, _ruleInfo1, null, null, false);
			break;
		}	
		case TOP_10:
		{
			final Double top10Val = _pickTop10(); // consider rank, bottom, 
			final RuleOperator opx = !bottom ? 
					RuleOperator.GREATER_THAN_OR_EQUAL:
					RuleOperator.LESS_THAN_OR_EQUAL;
			final CellRegion region = getRegions().iterator().next();
			final int row0 = region.row;
			final int col0 = region.column;
			final FormulaExpression expr = parseFormula(""+top10Val);
			_ruleInfo1 = new RuleInfo(getSheet(), this, expr, row0, col0);
			_matchable = new CellMatch2(opx, _ruleInfo1, null, null, false);
			break;
		}
		case DUPLICATE_VALUES:
		{
			Set<String> dupSet = _calcDuplicate();
			_matchable = new DuplicateCell(dupSet, false);
			break;
		}	
		case UNIQUE_VALUES:
		{
			Set<String> dupSet = _calcDuplicate();
			_matchable = new DuplicateCell(dupSet, true);
			break;
		}	
		case TIME_PERIOD:
			double[] res;
			switch(timePeriod) {
			case TOMORROW:
				res = DateUtil.calcTomorrow();
				break;
			case TODAY:
				res = DateUtil.calcToday();
				break;
			case YESTERDAY:
				res = DateUtil.calcYesterday();
				break;
			case LAST_7_DAYS:
				res = DateUtil.calcLast7Days();
				break;
			case NEXT_WEEK:
				res = DateUtil.calcNextWeek();
				break;
			case THIS_WEEK:
				res = DateUtil.calcThisWeek();
				break;
			case LAST_WEEK:
				res = DateUtil.calcLastWeek();
				break;
			case NEXT_MONTH:
				res = DateUtil.calcNextMonth();
				break;
			case THIS_MONTH:
				res = DateUtil.calcThisMonth();
				break;
			case LAST_MONTH:
				res = DateUtil.calcLastMonth();
				break;
			default:
				res = new double[] {0, 0};
			}
			_matchable = new DatesMatch2((int)res[0], (int)res[1]);
			break;
			
		case EXPRESSION:
		{
			final CellRegion region = getRegions().iterator().next();
			final int row0 = region.row;
			final int col0 = region.column;
			_ruleInfo1 = new RuleInfo(getSheet(), this, _formula1Expr, row0, col0);
			_matchable = new ExpressionMatch(_ruleInfo1);
			break;
		}	
		case COLOR_SCALE:
			if (_min == null) {
				vos = colorScale.getCFValueObjects();
				if (_prepareMinMax(vos) ) {
					_matchable = new NumMatch(); // true if a number
				} else {
					_matchable = new TrueMatch(false); // always false
				}
			}
			break;
		case DATA_BAR:
			if (_min == null) {
				vos = dataBar.getCFValueObjects();
				if (_prepareMinMax(vos)) {
					_matchable = new NumMatch(); // true if a number
				} else {
					_matchable = new TrueMatch(false); // always false
				}
			}
			break;
		case ICON_SET:
			if (_iconSetInfos == null) {
				vos = iconSet.getCFValueObjects();
				if (_prepareSteps(vos)) {
					_matchable = new NumMatch(); // true if a number
				} else {
					_matchable = new TrueMatch(false); // always false
				}
			}
			break;
		}
		return _matchable.match(cell);
	}

	//ZSS-1142: use to calculate min max for ColorScale and DataBar
	private boolean _prepareMinMax(List<SCFValueObject> vos) {
		if (vos != null) {
			_prepareVoHelper();
			final Double[] min_max = _voHelper.calcMinMax(vos);
			if (min_max == null) { // something wrong
				return false;
			} else {
				final int len = min_max.length;
				_min = min_max[0];
				switch(len) {
				case 2:
					_middle = null;
					_max = min_max[1];
					break;
					
				case 3:
					_middle = min_max[1];
					_max = min_max[2];
					break;
				}
				return true;
			}
		}
		return false;
	}

	//ZSS-1142: use to calculate steps for IconSet
	private boolean _prepareSteps(List<SCFValueObject> vos) {
		if (vos != null) {
			_prepareVoHelper();
			final Double[] min_max = _voHelper.calcMinMax(vos);
			if (min_max == null) { // something wrong
				return false;
			} else {
				final int len = min_max.length;
				_iconSetInfos = new IconSetInfo[len];
				for (int j = 0; j < len; ++j) {
					_iconSetInfos[j] = new IconSetInfo(min_max[j], vos.get(j).isGreaterOrEqual());
				}
				return true;
			}
		}
		return false;
	}

	//ZSS-1142
	private void _prepareVoHelper() {
		if (_voHelper == null) {
			_voHelper = new CFValueObjectHelper(this);
		}
	}
	
	//ZSS-1142: see https://en.wikipedia.org/wiki/Standard_deviation#Rapid_calculation_methods
	// return [average, stdv]
	private double[] _calcAverageStdv() {
		double a = 0;
		double q = 0;
		int k = 0;
		for (CellRegion region : this.getRegions()) {
			final int left = region.column;
			final int right = region.lastColumn;
			final int top = region.row;
			final int bottom = region.lastRow;
			for (int r = top; r <= bottom; ++r) {
				for (int col = left; col <= right; ++col) {
					final SCell cell = this.getSheet().getCell(r, col);
					final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
					if (cellval.getType() == CellType.NUMBER) {
						++k;
						final double x =  ((Number)cellval.getValue()).doubleValue();
						final double ak_1 = a;
						a = ak_1 + (x - ak_1) / k;
						q = q + (x - ak_1) * (x - a); 
					}
				}
			}
		}

		return new double[] {a, Math.sqrt(q/k)};
	}
	
	//ZSS-1142
	private Double _pickTop10() {
		if (rank <= 0) return null;
		boolean isTop = !bottom;
		boolean isPercent = percent;
		_prepareVoHelper();
		return _voHelper.pickTopX(isTop, isPercent, rank);
	}

	//ZSS-1142
	private Set<String> _calcDuplicate() {
		Map<CellValue, String> values = new HashMap<CellValue, String>();
		Set<String> keys = new HashSet<String>();
		for (CellRegion region : this.getRegions()) {
			final int left = region.column;
			final int right = region.lastColumn;
			final int top = region.row;
			final int bottom = region.lastRow;
			for (int row = top; row <= bottom; ++row) {
				for (int col = left; col <= right; ++col) {
					final SCell cell = this.getSheet().getCell(row, col);
					final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
					final String val = cellval == null ? null : values.get(cellval);
					if (val != null) {
						keys.add(val);
						keys.add(""+row+"_"+col);
					} else if (cellval != null && cellval.getType() != CellType.BLANK) {
						values.put(cellval, ""+row+"_"+col);
					}
				}
			}
		}
		return keys;
	}
	
	//ZSS-1142
	@Override
	public void destroy() {
		clearFormulaDependency(true);
		clearFormulaResultCache();
	}
	
	//ZSS-1142
	@Override
	public Collection<CellRegion> getRegions() {
		return _cfi.getRegions();
	}	

	//ZSS-1142
	@Override
	public boolean isFormulaParsingError() {
		boolean r = false;
		if(_formula1Expr!=null){
			r |= _formula1Expr.hasError();
		}
		if(!r && _formula2Expr!=null){
			r |= _formula2Expr.hasError();
		}
		if(!r && _formula3Expr!=null){
			r |= _formula3Expr.hasError();
		}
		return r;
	}

	//ZSS-1142
	@Override
	public String getFormula1() {
		return _unescapeFromPoi(_formula1Expr);
	}

	//ZSS-1142
	@Override
	public String getFormula2() {
		return _unescapeFromPoi(_formula2Expr);
	}

	//ZSS-1142
	@Override
	public String getFormula3() {
		return _unescapeFromPoi(_formula3Expr);
	}

	//ZSS-1251
	private void clearFormulaDependency(boolean all) {
		if(_formula1Expr!=null || _formula2Expr!=null || _formula3Expr!=null){
			Ref dependent = getRef();
			DependencyTable dt = 
			((AbstractBookSeriesAdv) getSheet().getBook().getBookSeries())
					.getDependencyTable();
			
			dt.clearDependents(dependent);
			
			// ZSS-648
			// must keep the region reference itself in DependencyTable; so add it back
			if (!all && this.getRegions() != null) {
				for (CellRegion regn : this.getRegions()) {
					dt.add(dependent, newDummyRef(getSheet().getSheetName(), regn));
				}
			}
		}
	}
	
	//ZSS-1251
	private Ref getRef(){
		return new ConditionalRefImpl(this._cfi);
	}
	
	//ZSS-1251
	public Ref getRef(String sheetName) {
		return new ConditionalRefImpl(getSheet().getBook().getBookName(), sheetName, _cfi.getId());
	}
	
	//ZSS-1251
	private Ref newDummyRef(String sheetName, CellRegion regn) {
		return new RefImpl(getSheet().getBook().getBookName(), sheetName, 
				regn.row, regn.column, regn.lastRow, regn.lastColumn);
	}

	//ZSS-1142
	@Override
	public void setFormula1(String formula1) {
		formula1 = _escapeToPoi(formula1);
		setEscapedFormulas(formula1, getEscapedFormula2(), getEscapedFormula3());
	}
	
	//ZSS-1142
	@Override
	public void setFormula2(String formula2) {
		formula2 = _escapeToPoi(formula2);
		setEscapedFormulas(getEscapedFormula1(), formula2, getEscapedFormula3());
	}

	//ZSS-1142
	@Override
	public void setFormula3(String formula3) {
		formula3 = _escapeToPoi(formula3);
		setEscapedFormulas(getEscapedFormula1(), getEscapedFormula2(), formula3);
	}

	//ZSS-1142
	private boolean isLiteralPtg(Ptg ptg) {
		return ptg instanceof BoolPtg
				|| ptg instanceof IntPtg 
				|| ptg instanceof NumberPtg
				|| ptg instanceof StringPtg
				|| ptg instanceof ErrPtg;
	}
	
	//ZSS-1142
	//20150108, henrichen: unescaped from POI before used in API.
	private String _unescapeFromPoi(FormulaExpression expr) { //ZSS-978, ZSS-994
		if (expr == null) return null;
		String formula = expr.getFormulaString();
		Ptg[] ptgs = expr.getPtgs();
		if (Strings.isBlank(formula)) return null;
		final StringBuilder sb = new StringBuilder();
		if (!formula.startsWith("\"") && formula.length() > 1) {
			if (ptgs.length > 1 || !isLiteralPtg(ptgs[0]))
				return sb.append("=").append(formula).toString(); //leading with '='  
		}
		if (ptgs.length == 1 && isLiteralPtg(ptgs[0])) {
			return ptgs[0] instanceof StringPtg ? ("=\""+formula+"\"") : ("="+formula);
		}
		return null;
	}
	
	//ZSS-1142
	//20150108, henrichen: formula must be escaped before store into POI
	private String _escapeToPoi(String formula) {
		if (Strings.isBlank(formula))
			return null;
		InputResult input = parseInput(formula);
		switch (input.getType()) {
		case FORMULA:
			return formula.substring(1);
		case STRING:
			return formula;
		case NUMBER:
			final Object val = input.getValue();
			double num = val instanceof Date ? EngineFactory.getInstance()
					.getCalendarUtil().dateToDoubleValue((Date) val)
					: ((Number) val).doubleValue();
					
			//ZSS-1147
			return getNumLocaleString(num);
		default:
			return formula;
		}
	}

	// ZSS-1142
	private InputResult parseInput(String formula) {
		final InputEngine ie = EngineFactory.getInstance().createInputEngine();
		return ie.parseInput(formula == null ? "" : formula,
				SCellStyle.FORMAT_GENERAL, new InputParseContext(ZssContext
						.getCurrent().getLocale()));
	}

	// ZSS-1142
	public FormulaExpression parseFormula(String formula) {
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		final Locale locale = ZssContext.getCurrent().getLocale();
		final SSheet sheet = getSheet();
		final FormulaParseContext formulaCtx = new FormulaParseContext(
				sheet.getBook(), sheet, null/* SCell */, sheet.getSheetName(),
				null, locale);
		final FormulaExpression expr = fe.parse(formula, formulaCtx);
		if (expr.hasError()) {
			String msg = expr.getErrorMessage();
			throw new InvalidFormulaException(msg == null ? "The formula ="
					+ formula + " contains error" : msg);
		}
		return expr;
	}
	
	//ZSS-1142
	@Override
	public void setFormulas(String formula1, String formula2, String formula3) {
		formula1 = _escapeToPoi(formula1);
		formula2 = _escapeToPoi(formula2);
		formula3 = _escapeToPoi(formula3);
		setEscapedFormulas(formula1, formula2, formula3);
	}

	//ZSS-1142
	@Override
	public void setEscapedFormulas(String formula1, String formula2, String formula3) {
		clearFormulaDependency(false); // will clear formula
		clearFormulaResultCache();

		if (formula1 != null) {
			_formula1Expr = parseFormula(formula1); // ZSS-994
		} else {
			_formula1Expr = null;
		}

		if (formula2 != null) {
			_formula2Expr = parseFormula(formula2); // ZSS-994
		} else {
			_formula2Expr = null;
		}
		if (formula3 != null) {
			_formula3Expr = parseFormula(formula3); // ZSS-994
		} else {
			_formula3Expr = null;
		}
	}

	//ZSS-1251
	//@since 3.9.0
	@Override
	public void copyFrom(SConditionalFormattingRule src, int rowOff, int colOff) {
		Validations.argInstance(src, ConditionalFormattingRuleImpl.class);
		ConditionalFormattingRuleImpl srcImpl = (ConditionalFormattingRuleImpl)src;

		SBook book = getSheet().getBook();
		type = srcImpl.type;
		operator = srcImpl.operator;
		if (srcImpl.style != null) {
			style = (SExtraStyle)((ExtraStyleImpl)srcImpl.style).cloneCellStyle(book);
		} else {
			style = null;
		}
		timePeriod = srcImpl.timePeriod;
		rank = srcImpl.rank;
		percent = srcImpl.percent;
		bottom = srcImpl.bottom;
		if (srcImpl.colorScale != null) {
			colorScale = ((ColorScaleImpl)srcImpl.colorScale).cloneColorScale(book);
		} else {
			colorScale = null;
		}
		if (srcImpl.dataBar != null) {
			dataBar = ((DataBarImpl)srcImpl.dataBar).cloneDataBar(book);
		} else {
			dataBar = null;
		}
		if (srcImpl.iconSet != null) {
			iconSet = ((IconSetImpl)srcImpl.iconSet).cloneIconSet();
		} else {
			iconSet = null;
		}
		text = srcImpl.text;
		notAboveAverage = srcImpl.notAboveAverage;
		equalAverage = srcImpl.equalAverage;
		standardDeviation = srcImpl.standardDeviation;

		shiftFormulas(srcImpl._formula1Expr, srcImpl._formula2Expr, srcImpl._formula3Expr, rowOff, colOff);
	}
	
	//ZSS-1142
	//@since 3.9.0
	/*package*/ ConditionalFormattingRuleImpl cloneConditionalFormattingRule(SConditionalFormatting cfi, SSheet sheet) {
		ConditionalFormattingRuleImpl tgt = new ConditionalFormattingRuleImpl(cfi); 
		SBook book = sheet.getBook();
		
		tgt.type = this.type;
		tgt.operator = this.operator;
		tgt.style = this.style;
		tgt.timePeriod = this.timePeriod;
		tgt.rank = this.rank;
		tgt.percent = this.percent;
		tgt.bottom = this.bottom;
		tgt.colorScale = ((ColorScaleImpl)this.colorScale).cloneColorScale(book);
		tgt.dataBar = ((DataBarImpl)this.dataBar).cloneDataBar(book);
		tgt.iconSet = ((IconSetImpl)this.iconSet).cloneIconSet();
		tgt.text = this.text;
		tgt.notAboveAverage = this.notAboveAverage;
		tgt.equalAverage = this.equalAverage;
		tgt.standardDeviation = this.standardDeviation;
				
		final String f1 = getEscapedFormula1();
		if (f1 != null) {
			final String f2 = getEscapedFormula2();
			final String f3 = getEscapedFormula3();
			tgt.setEscapedFormulas(f1, f2, f3); //ZSS-1183
		}

		return tgt;
	}
	
	//ZSS-1142
	/**
	 * 
	 * @param fe1
	 * @param fe2
	 * @since 3.9.0
	 */
	public void setFormulas(FormulaExpression fe1, FormulaExpression fe2, FormulaExpression fe3) {
		clearFormulaDependency(false); // will clear formula
		clearFormulaResultCache();
		
		_formula1Expr = fe1;
		_formula2Expr = fe2;
		_formula3Expr = fe3;
	}

	//ZSS-1142
	/**
	 * 
	 * @return
	 * @since 3.9.0
	 */
	public FormulaExpression getFormulaExpression1() {
		return _formula1Expr;
	}
	//ZSS-1142
	/**
	 * 
	 * @return
	 * @since 3.9.0
	 */
	public FormulaExpression getFormulaExpression2() {
		return _formula2Expr;
	}
	//ZSS-1142
	/**
	 * 
	 * @return
	 * @since 3.9.0
	 */
	public FormulaExpression getFormulaExpression3() {
		return _formula3Expr;
	}
	//ZSS-1142
	/**
	 * 
	 * @param formula
	 * @since 3.9.0
	 */
	public void setFormula1(FormulaExpression formula1) {
		setFormulas(formula1, _formula2Expr, _formula3Expr);
	}
	//ZSS-1142
	/**
	 * 
	 * @param formula
	 * @since 3.9.0
	 */
	public void setFormula2(FormulaExpression formula2) {
		setFormulas(_formula1Expr, formula2, _formula3Expr);
	}

	//ZSS-1142
	/**
	 * 
	 * @param formula
	 * @since 3.9.0
	 */
	public void setFormula3(FormulaExpression formula3) {
		setFormulas(_formula1Expr, _formula2Expr, formula3);
	}
	
	//ZSS-1142
	@Override
	public String getEscapedFormula1() {
		return _formula1Expr == null ? null : _formula1Expr.getFormulaString();
	}
	
	//ZSS-1142
	@Override
	public String getEscapedFormula2() {
		return _formula2Expr == null ? null : _formula2Expr.getFormulaString();
	}

	//ZSS-1142
	@Override
	public String getEscapedFormula3() {
		return _formula3Expr == null ? null : _formula3Expr.getFormulaString();
	}

	//ZSS-1142
	private String getNumLocaleString(double num) {
		final Locale locale = ZssContext.getCurrent().getLocale();
		final DataFormatter df= new DataFormatter(locale);
		final Format format = DataFormatter.isWholeNumber(num) ?
				df.getGeneralWholeNumJavaFormat() :
				df.getGeneralDecimalNumJavaFormat();
		return format.format(num);  
	}
	
	//ZSS-1142
	public SFill getColorScaleFill(Double value0) {
		if (_colorScaleCache == null) {
			_colorScaleCache = new HashMap<Double, SFill>();
		}
		SFill f = _colorScaleCache.get(value0);
		if (f != null) {
			return f;
		}
		f = _getColorScaleFill(value0.doubleValue());
		_colorScaleCache.put(value0, f);
		return f;
	}
	
	//ZSS-1142: return the X in X%.
	public double getDataBarPercent(Double value) {
		final double min = _min.doubleValue();
		final double max = _max.doubleValue();
		final int maxlen = dataBar.getMaxLength();
		final int minlen = dataBar.getMinLength();
		final int difflen = maxlen - minlen;
		final double diff0 = max - min;
		if (diff0 == 0.0) {
			return max == 0.0 ? minlen : maxlen;
		} else {
			if (value <= min) {
				return minlen;
			} else {
				return minlen + (value - min) / diff0 * difflen;
			}
		}
	}
	
	//ZSS-1142: return the iconSetId
	public int getIconSetId(Double value) {
		for (int j = _iconSetInfos.length; --j >= 0;) {
			final IconSetInfo info = _iconSetInfos[j]; 
			final double val = info.value;
			final boolean gte = info.gte;
			if (gte) {
				if (value >= val) return j;
			} else {
				if (value > val) return j;
			}
		}
		return 0;
	}

	//ZSS-1142
	private SFill _getColorScaleFill(double value) {
		SColor color = null;
		List<SColor> colors = colorScale.getColors(); 
		final SColor minColor = colors.get(0);
		final SColor maxColor = colors.get(colors.size()-1); 
		final double min = _min.doubleValue();
		final double max = _max.doubleValue();
		final double mid = _middle == null ? 0.0 : _middle.doubleValue();
		if (value >= max) {
			color = maxColor; // the last one
		} else if (value <= min) {
			color = minColor; // the first one
		} else { // in between
			if (_middle != null) {
				final SColor midColor = colors.get(1);
				if (value < mid) {
					//between min and middle
					color = _calcColor(value, min, mid, minColor, midColor);
				} else if (value > mid) {
					//between middle and max
					color = _calcColor(value, mid, max, midColor, maxColor);
				} else { 
					//== middle
					color = colors.get(1);
				}
			} else {
				//between min and max
				color = _calcColor(value, min, max, minColor, maxColor);
			}
		}
		return color == null ? null : new ExtraFillImpl(FillPattern.SOLID, color, color);
	}
	
	//ZSS-1142
	private SColor _calcColor(double value, double min, double max, SColor minColor, SColor maxColor) {
		final double diff0 = max - min;
		final double diff = value - min;
		final byte[] maxRgb = maxColor.getRGB();
		final byte[] minRgb = minColor.getRGB();
		final int r0 = minRgb[0] & 0xff;
		final int g0 = minRgb[1] & 0xff;
		final int b0 = minRgb[2] & 0xff;
		
		final int r1 = maxRgb[0] & 0xff;
		final int g1 = maxRgb[1] & 0xff;
		final int b1 = maxRgb[2] & 0xff;
		
		final int rDiff0 = r1 - r0;  
		final int gDiff0 = g1 - g0;  
		final int bDiff0 = b1 - b0;
		final int r = (int) Math.ceil(r0 + rDiff0 * diff / diff0); 
		final int g = (int) Math.ceil(g0 + gDiff0 * diff / diff0); 
		final int b = (int) Math.ceil(b0 + bDiff0 * diff / diff0);
		return new ColorImpl((byte)r, (byte)g, (byte)b);
	}
	
	//ZSS-1142
	private static class IconSetInfo {
		public final double value;
		public final boolean gte;
		IconSetInfo(double value, boolean gte) {
			this.value = value;
			this.gte = gte;
		}
	}
	

	//ZSS-1142
	@Override
	public void shiftFormulas(int rowOff, int colOff) {
		shiftFormulas(this._formula1Expr, this._formula2Expr, this._formula3Expr, rowOff, colOff);
	}
	
	//ZSS-1251
	private void shiftFormulas(FormulaExpression f1, FormulaExpression f2, FormulaExpression f3, int rowOff, int colOff) {
		if(f1!=null){
			FormulaExpression f1expr = null;
			FormulaExpression f2expr = null;
			FormulaExpression f3expr = null;
			if (rowOff != 0 || colOff != 0) {
				FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
				if(f1!=null){
					FormulaParseContext context = new FormulaParseContext(getSheet(), getSheet().getSheetName(), null); //nodependency
					f1expr = engine.shiftPtgs(f1, rowOff, colOff, context);//no dependency
				}
				
				if (f2 != null) {
					FormulaParseContext context = new FormulaParseContext(getSheet(), getSheet().getSheetName(), null); //nodependency
					f2expr = engine.shiftPtgs(f2, rowOff, colOff, context);//no dependency
				}
				
				if (f3 != null) {
					FormulaParseContext context = new FormulaParseContext(getSheet(), getSheet().getSheetName(), null); //nodependency
					f3expr = engine.shiftPtgs(f3, rowOff, colOff, context);//no dependency
				}
				setFormulas(f1expr, f2expr, f3expr);
			}
		}
	}

	//ZSS-1251
	public void addDependency(FormulaExpression fexpr) {
		final String bookName = getSheet().getBook().getBookName();
		final String sheetName = getSheet().getSheetName();
		final CellRegion region0 = getRegions().iterator().next();
		final int row0 = region0.row;
		final int col0 = region0.column;
		// row1Off, col1Off is used to move left-top point of formula areas
		// row2Off, col2Off is used to extends the formula area
		for (CellRegion region : getRegions()) {
			final int row1 = region.row;
			final int col1 = region.column;
			final int row1Off = row1 - row0;
			final int col1Off = col1 - col0;
			final int row2Off = region.getRowCount() - 1;
			final int col2Off = region.getColumnCount() - 1;
			for (Ptg ptg : fexpr.getPtgs()) {
				Ref precedent = toPrecedentRef(bookName, sheetName, ptg, row1Off, col1Off, row2Off, col2Off);
				if (precedent != null) {
					addDependency(precedent);
				}
			}
		}
	}

	//ZSS-1251
	private void addDependency(Ref precedent) {
		final SBookSeries bs = getSheet().getBook().getBookSeries();
		final DependencyTable dt = ((AbstractBookSeriesAdv)bs).getDependencyTable();
		dt.add(this.getRef(), precedent);
	}
	
	//ZSS-1251
	private Ref toPrecedentRef(String bookName, String sheetName, Ptg ptg, int row1Off, int col1Off, int row2Off, int col2Off) {
		if(ptg instanceof RefPtg) {
			final RefPtg rptg = (RefPtg)ptg;
			final boolean colRel = rptg.isColRelative();
			final boolean rowRel = rptg.isRowRelative();
			
			final int row1 = rptg.getRow() + (rowRel ? row1Off : 0);
			final int col1 = rptg.getColumn() + (colRel ? col1Off : 0);
			final int row2 = row1 + (rowRel ? row2Off : 0);
			final int col2 = col1 + (colRel ? col2Off : 0);
			return new RefImpl(bookName, sheetName, 
					Math.min(row1, row2), Math.min(col1, col2), 
					Math.max(row1, row2), Math.max(col1, col2));
		} else if(ptg instanceof AreaPtg) {
			AreaPtg aptg = (AreaPtg)ptg;
			final boolean col1Rel = aptg.isFirstColRelative();
			final boolean row1Rel = aptg.isFirstRowRelative();
			final boolean col2Rel = aptg.isLastColRelative();
			final boolean row2Rel = aptg.isLastRowRelative();
			final int row1 = aptg.getFirstRow() + (row1Rel ? row1Off : 0);
			final int col1 = aptg.getFirstColumn() + (col1Rel ? col1Off : 0);
			final int row2 = aptg.getLastRow() + (row2Rel ? row1Off + row2Off : 0);
			final int col2 = aptg.getLastColumn() + (col2Rel ? col1Off + col2Off : 0);
			return new RefImpl(bookName, sheetName,
					Math.min(row1, row2), Math.min(col1, col2), 
					Math.max(row1, row2), Math.max(col1, col2));
		}
		// TODO: should consider NamePtg...
		return null;
	}
	
	//ZSS-1142
	public RuleInfo getRuleInfo1() {
		return _ruleInfo1;
	}
	
	//ZSS-1142
	public RuleInfo getRuleInfo2() {
		return _ruleInfo2;
	}
}
