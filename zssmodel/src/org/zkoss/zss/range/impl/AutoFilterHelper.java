/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.range.impl;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TimeZone;

import org.zkoss.lang.Integers;
import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColorFilter;
import org.zkoss.zss.model.SCustomFilter;
import org.zkoss.zss.model.SCustomFilters;
import org.zkoss.zss.model.SDynamicFilter;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.STable;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.STop10Filter;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.model.impl.AbstractCellAdv;
import org.zkoss.zss.model.impl.CellStyleImpl;
import org.zkoss.zss.model.impl.CellValue;
import org.zkoss.zss.model.impl.ColorFilterImpl;
import org.zkoss.zss.model.impl.CustomFilterImpl;
import org.zkoss.zss.model.impl.CustomFiltersImpl;
import org.zkoss.zss.model.impl.ExtraFillImpl;
import org.zkoss.zss.model.impl.ExtraStyleImpl;
import org.zkoss.zss.model.impl.DynamicFilterImpl;
import org.zkoss.zss.model.impl.AbstractAutoFilterAdv.FilterColumnImpl;
import org.zkoss.zss.model.impl.Top10FilterImpl;
import org.zkoss.zss.model.util.CellStyleMatcher;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.range.impl.DataRegionHelper.FilterRegionHelper;
import org.zkoss.zss.range.impl.CellMatch;
import org.zkoss.zss.range.impl.Matchable;
import org.zkoss.zss.range.impl.GreaterThan;
import org.zkoss.zss.range.impl.GreaterThanOrEqual;
import org.zkoss.zss.range.impl.LessThan;
import org.zkoss.zss.range.impl.LessThanOrEqual;

/**
 * 
 * @author dennis
 *
 */
//these code if from XRangeImpl,XBookHelper and migrate to new model
/*package*/ class AutoFilterHelper extends RangeHelperBase{
	private static final long serialVersionUID = 764704415825488241L;
	private static final SCellStyle BLANK_STYLE = CellStyleImpl.BLANK_STYLE;
	private static Date MIN_DATE;
	static {
		Calendar cal = Calendar.getInstance();
		cal.set(0, 0, 0, 0, 0, 0);
		MIN_DATE = cal.getTime();
	}
	public AutoFilterHelper(SRange range){
		super(range);
	}
	
	public CellRegion findAutoFilterRegion() {
		return new DataRegionHelper(range).findAutoFilterDataRegion();
	}

	//ZSS-988
	public SAutoFilter enableTableFilter(STable table, final boolean enable){
		SAutoFilter filter = table.getAutoFilter();
		if(filter!=null && !enable){
			CellRegion region = filter.getRegion();
			SRange toUnhide = SRanges.range(sheet,region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn()).getRows();
			//to show all hidden row in autofiler region when disable
			toUnhide.setHidden(false);
			table.deleteAutoFilter();
			filter = null;
		}else if(filter==null && enable){
			table.enableAutoFilter(enable);
			filter = table.getAutoFilter();
		}
		return filter;
	}
	
	//refer to #XRangeImpl#autoFilter
	public SAutoFilter enableAutoFilter(final boolean enable){
		SAutoFilter filter = sheet.getAutoFilter();
		if(filter!=null && !enable){
			CellRegion region = filter.getRegion();
			SRange toUnhide = SRanges.range(sheet,region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn()).getRows();
			//to show all hidden row in autofiler region when disable
			toUnhide.setHidden(false);
			sheet.deleteAutoFilter();
			filter = null;
		}else if(filter==null && enable){
			CellRegion region = findAutoFilterRegion();
			if(region!=null){
				filter = sheet.createAutoFilter(region);
			}else{
				throw new InvalidModelOpException("can't find any data in range");
			}
		}
		return filter;
	}
	
	@Deprecated
	//refer to #XRangeImpl#autoFilter(int field, Object criteria1, int filterOp, Object criteria2, Boolean visibleDropDown) {
	public SAutoFilter enableAutoFilter(final int field, final FilterOp filterOp,
			final Object criteria1, final Object criteria2, final Boolean showButton) {
		STable table = ((AbstractSheetAdv)sheet).getTableByRowCol(getRow(), getColumn());
		return enableAutoFilter(table, field, filterOp, criteria1, criteria2, showButton);
	}
	
	//ZSS-988
	public SAutoFilter enableAutoFilter(STable table, final int field, final FilterOp filterOp,
			final Object criteria1, final Object criteria2, final Boolean showButton) {
		SAutoFilter filter = table == null ? sheet.getAutoFilter() : table.getAutoFilter();
		
		if(filter==null){
			//ZSS-988
			if (table != null) {
				table.enableAutoFilter(true);
				filter = table.getAutoFilter();
			} else {
				CellRegion region = new DataRegionHelper(range).findAutoFilterDataRegion();
				if(region!=null){
					filter = sheet.createAutoFilter(region);
				}else{
					throw new InvalidModelOpException("can't find any data in range");
				}
			}
		}
		enableAutoFilter0(table, filter, field, filterOp, criteria1, criteria2, showButton);
		return filter;
	}
	
	//ZSS-1193
	private Matchable<Double> getMatchByTop10Filter(STop10Filter top10Filter) {
		final boolean isTop = top10Filter.isTop();
		final Double filterVal = top10Filter.getFilterValue();
		return isTop ? 
				new GreaterThanOrEqual<Double>(filterVal) : 
				new LessThanOrEqual<Double>(filterVal);
	}
	
	//ZSS-1193
	private Matchable<Double> getMatchByDynamicFilter(SDynamicFilter dynaFilter) {
		final boolean isAbove = "aboveAverage".equals(dynaFilter.getType()); //ZSS-1234
		final Double avg = dynaFilter.getValue();
		return isAbove ? 
				new GreaterThan<Double>(avg) : 
				new LessThan<Double>(avg);
	}
	
	//ZSS-1234
	private Matchable<Date> getMatchDateByDynamicFilter(SDynamicFilter dynaFilter) {
		final FilterOp op = FilterOp.valueOf(dynaFilter.getType());
		switch(op) {
		case tomorrow:
		case today:
		case yesterday:
		case nextWeek:
		case thisWeek:
		case lastWeek:
		case nextMonth:
		case thisMonth:
		case lastMonth:
		case nextQuarter:
		case thisQuarter:
		case lastQuarter:
		case nextYear:
		case thisYear:
		case lastYear:
		case yearToDate:
		{
			final Double min = dynaFilter.getValue();
			final Double max = dynaFilter.getMaxValue();
			return new DatesMatch(min.intValue(), max.intValue());
		}
		case Q1:
			return new QuarterMatch(0, 3);
		case Q2:
			return new QuarterMatch(3, 6);
		case Q3:
			return new QuarterMatch(6, 9);
		case Q4:
			return new QuarterMatch(9, 12);
		case M1:
			return new MonthMatch(0);
		case M2:
			return new MonthMatch(1);
		case M3:
			return new MonthMatch(2);
		case M4:
			return new MonthMatch(3);
		case M5:
			return new MonthMatch(4);
		case M6:
			return new MonthMatch(5);
		case M7:
			return new MonthMatch(6);
		case M8:
			return new MonthMatch(7);
		case M9:
			return new MonthMatch(8);
		case M10:
			return new MonthMatch(9);
		case M11:
			return new MonthMatch(10);
		case M12:
			return new MonthMatch(11);
		}
		return null;
	}
	
	//ZSS-1193
	private Double _pickTop10(int col, int row, int row2, int value, boolean isTop, boolean isPercent) {
		if (value <= 0) return null;
		
		int count = 0;
		boolean error = false;
		List<Double> list = new ArrayList<Double>();
		for (int r = row+1; r <= row2; ++r) { // should exclude the cell with the dropdown button
			final SCell cell = sheet.getCell(r, col);
			final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
			if (cellval.getType() == CellType.ERROR) {
				error = true;
				break;
			}
			if (cellval.getType() == CellType.NUMBER) {
				list.add(((Number)cellval.getValue()).doubleValue());
				++count;
			}
		}
		
		if (error) return null; // no filtering
		
		Collections.sort(list);
		if (isPercent) {
			count = Math.min((int) Math.ceil(count * value / 100.0), count); // 0.+ should be 1
		} else {
			count = Math.min(count, value);
		}
		
		return isTop ? list.get(list.size()-count) : list.get(count-1);
	}
	

	//ZSS-1193
	private Double _calcAverage(int col, int row, int row2) {
		double total = 0;
		int count = 0;
		boolean error = false;
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col);
			final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
			if (cellval.getType() == CellType.ERROR) {
				error = true;
				break;
			}
			if (cellval.getType() == CellType.NUMBER) {
				total += ((Number)cellval.getValue()).doubleValue();
				++count;
			}
		}
		
		if (error) return null; // no filtering
		return Double.valueOf(total / count);
	}
	
	//ZSS-1193: Top10/Top10Percent/Bottom10/Bottom10Percent
	private LinkedHashMap<Integer, Boolean> _filterByTop10Filter(SAutoFilter filter, NFilterColumn fc, int field) {
		final STop10Filter top10Filter = fc.getTop10Filter();
		final Matchable<Double> match = getMatchByTop10Filter(top10Filter);
		return _filterByNumber(filter, fc, field, match);
	}
	
	//ZSS-1193: AboveAverage/BelowAverage
	private LinkedHashMap<Integer, Boolean> _filterByDynamicFilter(SAutoFilter filter, NFilterColumn fc, int field) {
		final SDynamicFilter dynaFilter = fc.getDynamicFilter();
		SAutoFilter.FilterOp op = SAutoFilter.FilterOp.valueOf(dynaFilter.getType());
		switch(op) {
		case aboveAverage:
		case belowAverage:
			final Matchable<Double> match = getMatchByDynamicFilter(dynaFilter);
			return _filterByNumber(filter, fc, field, match);
		default:
			//ZSS-1234
			final Matchable<Date> matchDate = getMatchDateByDynamicFilter(dynaFilter);
			return _filterByDate(filter, fc, field, matchDate);
		}
	}
		
	private LinkedHashMap<Integer, Boolean> _filterByNumber(SAutoFilter filter, NFilterColumn fc, int field, Matchable<Double> match) {
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();
		
		LinkedHashMap<Integer, Boolean> affectedRows = new LinkedHashMap<Integer, Boolean>();
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col);
			final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
			Object val = cellval.getValue();
			//ZSS-1234
			if (val instanceof Date) {
				val = DateUtil.getExcelDate((Date)val);
			}
			if (cellval.getType() == CellType.NUMBER
				&& match.match((Double)val)) {
				 //candidate to be shown (other fildColumn might still hide this row!
				final SRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					affectedRows.put(r, false);
				}
			} else { //to be hidden
				final SRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					affectedRows.put(r, true);
				}
			}
		}
		return affectedRows;
	}

	//ZSS-1234
	private LinkedHashMap<Integer, Boolean> _filterByDate(SAutoFilter filter, NFilterColumn fc, int field, Matchable<Date> match) {
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();
		
		LinkedHashMap<Integer, Boolean> affectedRows = new LinkedHashMap<Integer, Boolean>();
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col);
			final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
			Object val = cellval.getValue();
			if (cellval.getType() == CellType.NUMBER && !(val instanceof Date)) {
				val = DateUtil.getJavaDate((Double) val, TimeZone.getTimeZone("UTC"));
			}
			
			if (cellval.getType() == CellType.NUMBER
				&& match.match((Date)val)) {
				 //candidate to be shown (other fildColumn might still hide this row!
				final SRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					affectedRows.put(r, false);
				}
			} else { //to be hidden
				final SRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					affectedRows.put(r, true);
				}
			}
		}
		return affectedRows;
	}

	//ZSS-1193
	// criteria1: [count, isTop, isPercent]
	private STop10Filter _prepareTop10Filter(Object[] criteria, CellRegion region, int field, FilterOp op) {
		final int row = region.getRow();
		final int row2 = region.getLastRow();
		final int col = region.getColumn() + field - 1;
		final boolean isTop = criteria[1] == Boolean.TRUE; // top or bottom
		final boolean isPercent = criteria[2] == Boolean.TRUE; // percent or value
		final int value = ((Integer)criteria[0]).intValue();
		final Double filterVal = _pickTop10(col, row, row2, value, isTop, isPercent);
		return filterVal == null ? null : new Top10FilterImpl(isTop, value, isPercent, filterVal);
	}
	
	//ZSS-1193, ZSS-1234
	private SDynamicFilter _prepareDynamicFilter(CellRegion region, int field, FilterOp op) {
		final int row = region.getRow();
		final int row2 = region.getLastRow();
		final int col = region.getColumn() + field - 1;
		Double maxVal = null;
		Double val = null;
		boolean fail = false;
		switch(op) {
		case aboveAverage:
			val = _calcAverage(col, row, row2);
			if (val == null) {
				fail = true;
			}
			break;
			
		case belowAverage:	
			val = _calcAverage(col, row, row2);
			if (val == null) {
				fail = true;
			}
			break;

		//ZSS-1234
		case tomorrow:
		{
			double[] res = DateUtil.calcTomorrow(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case today:
		{
			double[] res = DateUtil.calcToday(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}	
		case yesterday:
		{
			double[] res = DateUtil.calcYesterday(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
			
		case nextWeek:
		{
			double[] res = DateUtil.calcNextWeek(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}	
		case thisWeek:
		{
			double[] res = DateUtil.calcThisWeek(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case lastWeek:
		{
			double[] res = DateUtil.calcLastWeek(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case nextMonth:
		{
			double[] res = DateUtil.calcNextMonth(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}	
		case thisMonth:
		{
			double[] res = DateUtil.calcThisMonth(); //ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case lastMonth:
		{
			double[] res = DateUtil.calcLastMonth();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case nextQuarter:
		{
			double[] res = DateUtil.calcNextQuarter();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}	
		case thisQuarter:
		{
			double[] res = DateUtil.calcThisQuarter();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case lastQuarter:
		{
			double[] res = DateUtil.calcLastQuarter();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case nextYear:
		{
			double[] res = DateUtil.calcNextYear();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}	
		case thisYear:
		{
			double[] res = DateUtil.calcThisYear();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case lastYear:
		{
			double[] res = DateUtil.calcLastYear();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		case yearToDate:
		{
			double[] res = DateUtil.calcYearToDate();//ZSS-1142
			val = res[0];
			maxVal = res[1];
			break;
		}
		}
		return fail ? null : new DynamicFilterImpl(maxVal, val, op.name()); //ZSS-1234
	}

	//ZSS-1192
	// criteria: [SCustomFitler.Operator1, val1, SCustomFilter.Operator2, val2]
	private SCustomFilters _prepareCustomFilters(String[] criteria1, String[] criteria2, boolean isAnd) {
		final FilterOp op1 = FilterOp.valueOf(criteria1[0]);
		final String val1 = criteria1[1];
		final SCustomFilter f1 = new CustomFilterImpl(val1, op1);
		
		final FilterOp op2 = criteria2 != null ? 
				FilterOp.valueOf(criteria2[0]) : null;
		final String val2 = criteria2 != null ? criteria2[1] : null;
		final SCustomFilter f2 = op2 == null ? null : new CustomFilterImpl(val2, op2);
		
		return new CustomFiltersImpl(f1, f2, isAnd);
	}
	
	//ZSS-1191
	// criteria[0]: pattern name(see FillPattern); [1]: foreground color; [2]: background color 
	private SColorFilter _prepareColorFilter(String[] criteria, boolean byFontColor) {
		final FillPattern pattern = FillPattern.valueOf(criteria[0]);
		final String fg = criteria[1];
		final String bg = criteria[2];
		
		final SFill fill = new ExtraFillImpl(pattern, fg, bg);
		final SExtraStyle src = new ExtraStyleImpl(null, fill, null, null);
		final CellStyleMatcher matcher = new CellStyleMatcher(src);
		final SBook book = range.getSheet().getBook();
		
		SExtraStyle style = book.searchExtraStyle(matcher);
		if (style == null) {
			book.addExtraStyle(src);
			style = src;
		}
		return new ColorFilterImpl(style, byFontColor);
	}

	//ZSS-1192
	private Matchable<SCell> getMatchByCustomFilters(SCustomFilters filters, int filterType) {
		final SCustomFilter f1 = filters.getCustomFilter1();
		final SCustomFilter f2 = filters.getCustomFilter2();
		final boolean isAnd = filters.isAnd();
		
		return new CellMatch(f1, f2, isAnd);
				
	}
	
	//ZSS-1192
	private LinkedHashMap<Integer, Boolean> _filterByCustomFilters(SAutoFilter filter, NFilterColumn fc, int field) {
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();

		final SCustomFilters filters = fc.getCustomFilters();
		final int filterType = ((FilterColumnImpl)fc).getFilterType();
		
		final Matchable<SCell> match = getMatchByCustomFilters(filters, filterType);
		
		LinkedHashMap<Integer, Boolean> affectedRows = new LinkedHashMap<Integer, Boolean>(); 
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col);
			if (!match.match(cell.isNull() ? null : cell)) { //to be hidden
				final SRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					affectedRows.put(r, true);
				}
			} else { //candidate to be shown (other FieldColumn might still hide this row!
				final SRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					affectedRows.put(r, false);
				}
			}
		}
		return affectedRows;
	}
	
	//ZSS-1191
	private boolean _match(SCellStyle style, SFill fill, boolean byFontColor) {
		if (byFontColor) {
			return fill.getFillColor().equals(style.getFont().getColor());
		} else {
			return fill.equals(style.getFill());
		}
	}
	
	//ZSS-1191
	private LinkedHashMap<Integer, Boolean> _filterByColor(SAutoFilter filter, NFilterColumn fc, int field) {
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();

		SFill fill = fc.getColorFilter().getExtraStyle().getFill();
		final boolean byFontColor = fc.getColorFilter().isByFontColor();
		
		if (!byFontColor && fill.getFillPattern() == FillPattern.SOLID) {
			fill = new ExtraFillImpl(FillPattern.SOLID, fill.getBackColor(), fill.getFillColor());
		}
		
		LinkedHashMap<Integer, Boolean> affectedRows = new LinkedHashMap<Integer, Boolean>(); 
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col);
			final SCellStyle style = cell.isNull() ? null : cell.getCellStyle();
			if (!_match(style, fill, byFontColor)) { //to be hidden
				final SRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					affectedRows.put(r, true);
				}
			} else { //candidate to be shown (other FieldColumn might still hide this row!
				final SRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					affectedRows.put(r, false);
				}
			}
		}
		return affectedRows;
	}
	
	//ZSS-1191
	LinkedHashMap<Integer, Boolean> _filterByValues(SAutoFilter filter, NFilterColumn fc, int field) {
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();

		final Set cr1 = fc.getCriteria1();
		//ZSS-1083(refix ZSS-838): Collect affected rows first
		LinkedHashMap<Integer, Boolean> affectedRows = new LinkedHashMap<Integer, Boolean>(); 
//		final Set<Ref> all = new HashSet<Ref>();
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col); 
			final String val = isBlank(cell) ? "=" : getFormattedText(cell); //"=" means blank!
			if (cr1 != null && !cr1.isEmpty() && !cr1.contains(val)) { //to be hidden
				final SRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					//ZSS-1083(refix ZSS-838): Collect affected rows first 
//					SRanges.range(sheet,r,0).getRows().setHidden(true);
					affectedRows.put(r, true);
				}
			} else { //candidate to be shown (other FieldColumn might still hide this row!
				final SRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					// ZSS-646: we don't care about the columns at all; use 0.
//					final int left = sheet.getStartCellIndex(r);
//					final int right = sheet.getEndCellIndex(r);
					//ZSS-1083(refix ZSS-838): Collect affected rows first
//					final SRange rng = SRanges.range(sheet,r,0,r,0);  
//					all.addAll(rng.getRefs());
//					rng.getRows().setHidden(false); //unhide row
					affectedRows.put(r, false);
					
//					rng.notifyChange(); //why? text overflow? ->  //BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
				}
			}
		}
		return affectedRows;
	}
	
	//ZSS-988
	private void enableAutoFilter0(STable table, SAutoFilter filter, final int field, final FilterOp filterOp,
			Object criteria1, final Object criteria2, final Boolean showButton) {
		final NFilterColumn fc = filter.getFilterColumn(field-1,true);
		
		//ZSS-1191
		Map<String, Object> extra = new HashMap<String, Object>();
		
		SDynamicFilter dynaFilter = null; //ZSS-1193
		STop10Filter top10Filter = null; //ZSS-1193
		switch(filterOp) {
		//ZSS-1191
		case cellColor:
		case fontColor:
			extra.put("colorFilter", _prepareColorFilter((String[])criteria1, filterOp == FilterOp.fontColor));
			criteria1 = null;
			break;
		
		//ZSS-1192
		case and:
		case or:
			extra.put("customFilters", _prepareCustomFilters((String[])criteria1, (String[])criteria2, filterOp == FilterOp.and));
			criteria1 = null;
			break;
		
		//ZSS-1193: SDynamicFilter
		case aboveAverage:
		case belowAverage:
			
		//ZSS-1234: SDynamicFilter
		case tomorrow:
		case today:
		case yesterday:
		case nextWeek:
		case thisWeek:
		case lastWeek:
		case nextMonth:
		case thisMonth:
		case lastMonth:
		case nextQuarter:
		case thisQuarter:
		case lastQuarter:
		case nextYear:
		case thisYear:
		case lastYear:
		case yearToDate:
		case Q1:
		case Q2:
		case Q3:
		case Q4:
		case M1:
		case M2:
		case M3:
		case M4:
		case M5:
		case M6:
		case M7:
		case M8:
		case M9:
		case M10:
		case M11:
		case M12:
			//if #Error in row; the dynamic filter will be ignored
			dynaFilter = 
				_prepareDynamicFilter(filter.getRegion(), field, filterOp); //ZSS-1234 
			extra.put("dynamicFilter", dynaFilter == null ? DynamicFilterImpl.NOOP_DYNAFILTER : dynaFilter);
			criteria1 = null;
			break;

		//ZSS-1193: STop10Filter
		case top10:
			//if #Error in row; the top10 filter will be ignored
			top10Filter = 
				_prepareTop10Filter((Object[])criteria1, filter.getRegion(), field, filterOp);
			extra.put("top10Filter", top10Filter == null ? Top10FilterImpl.NOOP_TOP10FILTER : top10Filter);
			criteria1 = null;
			break;
		}
		
		fc.setProperties(filterOp, criteria1, criteria2, showButton, extra);

		//update rows
		LinkedHashMap<Integer, Boolean> affectedRows = null;
		switch(filterOp) {
		//ZSS-1191
		case cellColor: 
		case fontColor:
			affectedRows = _filterByColor(filter, fc, field);
			break;
			
		//ZSS-1192
		case or:
		case and:
			affectedRows = _filterByCustomFilters(filter, fc, field);
			break;
			
		//ZSS-1193
		case aboveAverage:
		case belowAverage:
			
		//ZSS-1234
		case tomorrow:
		case today:
		case yesterday:
		case nextWeek:
		case thisWeek:
		case lastWeek:
		case nextMonth:
		case thisMonth:
		case lastMonth:
		case nextQuarter:
		case thisQuarter:
		case lastQuarter:
		case nextYear:
		case thisYear:
		case lastYear:
		case yearToDate:
		case Q1:
		case Q2:
		case Q3:
		case Q4:
		case M1:
		case M2:
		case M3:
		case M4:
		case M5:
		case M6:
		case M7:
		case M8:
		case M9:
		case M10:
		case M11:
		case M12:
			if (dynaFilter != null) {
				affectedRows = _filterByDynamicFilter(filter, fc, field);
			}
			break;
			
		//ZSS-1193
		case top10:
			if (top10Filter != null) {
				affectedRows = _filterByTop10Filter(filter, fc, field);
			}
			break;
			
		default:
			affectedRows = _filterByValues(filter, fc, field);
		}
						
		//ZSS-1083(refix ZSS-838): Handle affected rows
		if (affectedRows != null && !affectedRows.isEmpty()) {
			final String key = (table == null ? sheet.getId() : table.getName())+"_ZSS_AFFECTED_ROWS"; 
			Executions.getCurrent().setAttribute("CONTAINS_"+key, true);
			int sz = affectedRows.size();
			int j  = 0;
			for (int r : affectedRows.keySet()) {
				//ZSS-838: flag only the last handled row so 
				//  Spreadsheet.java#updateAutoFilter can optimize the smartUpdate
				if (++j == sz) { 
					Executions.getCurrent().setAttribute(key, new Integer(sz));
				} else { // wait for last affected row
					Executions.getCurrent().setAttribute(key, Integers.ZERO);
				}
				SRanges.range(sheet,r,0).getRows().setHidden(affectedRows.get(r));
			}
		}
//		BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
	}
	
	private boolean canUnhide(SAutoFilter af, NFilterColumn fc, int row, int col) {
		final Collection<NFilterColumn> fltcs = af.getFilterColumns();
		for(NFilterColumn fltc: fltcs) {
			if (fc.equals(fltc)) continue;
			if (shallHide(fltc, row, col)) { //any FilterColumn that shall hide the row
				return false;
			}
		}
		return true;
	}
	
	private boolean shallHide(NFilterColumn fc, int row, int col) {
		final SCell cell = sheet.getCell(row, col + fc.getIndex());

		//ZSS-1191
		SColorFilter colorFilter = fc.getColorFilter();
		if (colorFilter != null) {
			SCellStyle style = cell.isNull() ? BLANK_STYLE : cell.getCellStyle();
			return !_match(style, colorFilter.getExtraStyle().getFill(), colorFilter.isByFontColor());
		}
		
		//ZSS-1192
		SCustomFilters custFilters = fc.getCustomFilters();
		final int filterType = ((FilterColumnImpl)fc).getFilterType();
		if (custFilters != null) {
			Matchable<SCell> match = getMatchByCustomFilters(custFilters, filterType);
			return !match.match(cell.isNull() ? null : cell);
		}
		
		//ZSS-1193, ZSS-1234
		SDynamicFilter dynaFilter = fc.getDynamicFilter();
		if (dynaFilter != null) {
			final String type = dynaFilter.getType();
			if ("aboveAverage".equals(type) || "belowAverage".equals(type)) { 
				//ZSS-1193
				Matchable<Double> match = getMatchByDynamicFilter(dynaFilter);
				CellValue cv = ((AbstractCellAdv)cell).getEvalCellValue(true);
				return cv.getType() != CellType.NUMBER || !match.match((Double)cv.getValue());
			} else { 
				//ZSS-1234
				Matchable<Date> match = getMatchDateByDynamicFilter(dynaFilter);
				CellValue cv = ((AbstractCellAdv)cell).getEvalCellValue(true);
				Object val = cv.getValue();
				if (cv.getType() == CellType.NUMBER && !(val instanceof Date)) {
					val = DateUtil.getJavaDate((Double)val, TimeZone.getTimeZone("UTC"));
				}
				return cv.getType() != CellType.NUMBER || !match.match((Date)val);
			}
		}
		
		//ZSS-1193
		STop10Filter top10Filter = fc.getTop10Filter();
		if (top10Filter != null) {
			Matchable<Double> match = getMatchByTop10Filter(top10Filter);
			CellValue cv = ((AbstractCellAdv)cell).getEvalCellValue(true);
			return cv.getType() != CellType.NUMBER || !match.match((Double)cv.getValue());
		}
		
		final boolean blank = isBlank(cell); 
		final String val =  blank ? "=" : getFormattedText(cell); //"=" means blank!
		final Set critera1 = fc.getCriteria1();
		return critera1 != null && !critera1.isEmpty() && !critera1.contains(val);
	}
	
	@Deprecated
	//refer to XRangeImpl#showAllData
	public void resetAutoFilter() {
		//ZSS-988
		STable table = ((AbstractSheetAdv)sheet).getTableByRowCol(getRow(), getColumn());
		resetAutoFilter(table);
	}
	//ZSS-988: check if this filter ever filter out any rows; so it can do
	// resetAutoFilter() or reapplyAutoFilter()
	//@since 3.8.0
	private void validFiltered(SAutoFilter af) {
		if (af == null) { //no AutoFilter to apply 
			return;
		}
		final Collection<NFilterColumn> fcs = af.getFilterColumns();
		if (fcs == null)
			return;
		
		//ZSS-988: must contains filterColumn with criteria that can be cleared
		boolean hasCriteria1 = false;
		for(NFilterColumn fc : fcs) {
			final Set criteria1 = fc.getCriteria1();
			if (criteria1 != null && !criteria1.isEmpty()) {
				hasCriteria1 = true;
				break;
			}
		}
		if (!hasCriteria1) {
			throw new InvalidModelOpException("The filter is not applied any criteria"); 
		}
	}
	
	public void resetAutoFilter(STable table) {
		final SAutoFilter af = table == null ? sheet.getAutoFilter() : table.getAutoFilter();
		if (af == null) { //no AutoFilter to apply 
			return;
		}
		final CellRegion afrng = af.getRegion();
		final Collection<NFilterColumn> fcs = af.getFilterColumns();
		if (fcs == null)
			return;
		
		//ZSS-988: filterColumn been filtering with criteria that can be cleared
		validFiltered(af);
		
		for(NFilterColumn fc : fcs) {
			fc.setProperties(FilterOp.values, null, null, null); //clear all filter
		}
		final int row1 = afrng.getRow();
		final int row = row1 + 1;
		final int row2 = afrng.getLastRow();
		final int col1 = afrng.getColumn();
		final int col2 = afrng.getLastColumn();
//		final Set<Ref> all = new HashSet<Ref>();
		for (int r = row; r <= row2; ++r) {
			final SRow rowobj = sheet.getRow(r);
			if (rowobj.isHidden()) { //a hidden row
				//ZSS-646, we don't care about columns, use 0.
				//final int left = sheet.getStartCellIndex(r);
				//final int right = sheet.getEndCellIndex(r);
				final SRange rng = SRanges.range(sheet,r,0,r,0); 
//				all.addAll(rng.getRefs());
				rng.getRows().setHidden(false); //unhide
			}
		}

//		BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
		//update button
//		final XRangeImpl buttonChange = (XRangeImpl) XRanges.range(_sheet, row1, col1, row1, col2);
//		BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
	}

	@Deprecated
	//refer to XRangeImpl#applyFilter
	public void applyAutoFilter() {
		//ZSS-988
		STable table = ((AbstractSheetAdv)sheet).getTableByRowCol(getRow(), getColumn());
		applyAutoFilter(table);
	}
	
	//ZSS-988
	//@since 3.8.0
	public void applyAutoFilter(STable table) {
		final SAutoFilter oldFilter = table == null ? sheet.getAutoFilter() : table.getAutoFilter();
		
		if (oldFilter==null) { //no criteria is applied
			return;
		}

		//ZSS-988: filterColumn been filtering with criteria that can be reapplied
		validFiltered(oldFilter);

		CellRegion region = oldFilter.getRegion();
		//copy filtering criteria
		int firstRow = region.getRow(); //first row is header
		int firstColumn = region.getColumn();
		//backup column data because getting from removed auto filter will cause XmlValueDisconnectedException
		
		//index,criteria1,op,criteria2,showVisible
		List<Object[]> originalFilteringColumns = new ArrayList<Object[]>();
		if (oldFilter.getFilterColumns() != null){ //has applied some criteria
			for (NFilterColumn filterColumn : oldFilter.getFilterColumns()){
				Object[] filterColumnData = new Object[5];
				filterColumnData[0] = filterColumn.getIndex()+1;
				filterColumnData[1] = filterColumn.getCriteria1().toArray(new String[0]);
				filterColumnData[2] = filterColumn.getOperator();
				filterColumnData[3] = filterColumn.getCriteria2();
				filterColumnData[4] = filterColumn.isShowButton();
				originalFilteringColumns.add(filterColumnData);
			}
		}
		
		SAutoFilter newFilter = null;
		//ZSS-988
		if (table != null) {
			enableTableFilter(table, false); // unhidden rows if any
			newFilter = enableTableFilter(table, true); // create a new filter
		} else {
			enableAutoFilter(false); //disable existing filter
			//re-define filtering range 
			CellRegion filteringRange = new FilterRegionHelper().findCurrentRegion(sheet, firstRow, firstColumn);
			if (filteringRange == null){ //Don't enable auto filter if there are all blank cells
				return;
			}else{
				//enable auto filter
				newFilter = sheet.createAutoFilter(filteringRange);
				//			BookHelper.notifyAutoFilterChange(getRefs().iterator().next(),true);
			}
		}
		
		//apply original criteria
		for (int nCol = 0 ; nCol < originalFilteringColumns.size(); nCol ++){
			Object[] oldFilterColumn =  originalFilteringColumns.get(nCol);
			
			int field = (Integer)oldFilterColumn[0];
			Object c1 = oldFilterColumn[1];
			FilterOp op = (FilterOp)oldFilterColumn[2];
			Object c2 = oldFilterColumn[3];
			boolean showBtn = (Boolean)oldFilterColumn[4];
			
			enableAutoFilter0(table, newFilter, field,op,c1,c2,showBtn); //ZSS-988
		}
	}

}
