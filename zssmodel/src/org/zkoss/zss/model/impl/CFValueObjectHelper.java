/* CFValueObjectHelper.java

	Purpose:
		
	Description:
		
	History:
		Jun 30, 2016 10:42:20 AM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;

//ZSS-1142
/**
 * Helper on calculating SCFValueObject.
 * @author Henri
 * @since 3.9.0
 */
public class CFValueObjectHelper {
	private ConditionalFormattingRuleImpl rule;
	private ArrayList<Double> _sortedList; // number sorted into list
	
	public CFValueObjectHelper(ConditionalFormattingRuleImpl rule) {
		this.rule = rule;
	}

	//ZSS-1142
	public Double[] calcMinMax(List<SCFValueObject> vos) {
		Double[] values = new Double[vos.size()];
		for (int j = 0, len = vos.size(); j < len; ++j) {
			SCFValueObject vo = vos.get(j);
			Double val = _calcVo(vo);
			if (val == null) {
				return null;
			}
			values[j] = val;
		}
		return values;
	}
	
	private Double _calcVo(SCFValueObject vo) {
		switch(vo.getType()) {
		case MIN:
		{
			_prepareSortedList();
			return _sortedList.isEmpty() ? null : _sortedList.get(0); 
		}
			
		case MAX:
		{
			_prepareSortedList();
			return _sortedList.isEmpty() ? null : _sortedList.get(_sortedList.size() - 1);
		}
			
		case FORMULA:
		{
			final CellRegion region = rule.getRegions().iterator().next();
			final int row0 = region.row;
			final int col0 = region.column;
			final String formula = vo.getValue();
			final FormulaExpression expr = rule.parseFormula(formula);
			((CFValueObjectImpl)vo).setFormulaExpression(expr); //ZSS-1151
			final RuleInfo rulex = new RuleInfo(rule.getSheet(), rule, expr, row0, col0);
			final EvaluationResult result = rulex.evalFormula(null);
			if(result.getType() == ResultType.SUCCESS){
				final Object val = result.getValue();
				if (val instanceof Double) {
					return (Double)val;
				}
				if (val instanceof Date) {
					return Double.valueOf(DateUtil.getExcelDate((Date)val));
				}
			}
			//string, blank, error, boolean is deemed as false
			return null;
		}	
		case NUM:
		{
			try {
				return Double.valueOf(vo.getValue());
			} catch (NumberFormatException ex) {
				return null;
			}
		}
		
		case PERCENT:
		{
			try {
				final double rank0 = Double.valueOf(vo.getValue());
				return _pickPercent(rank0);
			} catch (NumberFormatException ex) {
				return null;
			}
		}
			
		case PERCENTILE:
		{
			try {
				final double rank0 = Double.valueOf(vo.getValue());
				return _pickPercentileInc(rank0);
			} catch (NumberFormatException ex) {
				return null;
			}
		}
		}
		return null;
	}
	
	private void _prepareSortedList() {
		if (_sortedList == null) {
			_sortedList = new ArrayList<Double>();
			for (CellRegion rgn : rule.getRegions()) {
				final int l = rgn.column;
				final int t = rgn.row;
				final int r0 = rgn.lastColumn;
				final int b = rgn.lastRow;
			
				for (int row = t; row <= b; ++row) { // should exclude the cell with the dropdown button
					for (int col = l; col <= r0; ++col) {
						final SCell cell = rule.getSheet().getCell(row, col);
						final CellValue cellval = ((AbstractCellAdv)cell).getEvalCellValue(true);
						if (cellval.getType() == CellType.NUMBER) {
							_sortedList.add(((Number)cellval.getValue()).doubleValue());
						}
					}
				}
			}
			Collections.sort(_sortedList);
		}
	}

	public Double pickTopX(boolean isTop, boolean isPercent, double rank0) {
		_prepareSortedList();
		int count = _sortedList.size();
		if (count == 0) { //nothing to pick
			return null;
		}
		if (isPercent) {
			return _pickPercentileInc(isTop ? 100 - rank0 : rank0);
// see https://en.wikipedia.org/wiki/Percentile#The_Nearest_Rank_method
//			count = Math.min((int) Math.ceil(count * rank0 / 100.0), count); // 0.+ should be 1
		} else {
			count = Math.min(count, (int)rank0);
			return isTop ? _sortedList.get(_sortedList.size()-count) : _sortedList.get(count == 0 ? 0 : count-1);
		}
		
	}
	
	// http://www.excel-easy.com/examples/icon-sets.html
	private Double _pickPercent(double rank0) {
		_prepareSortedList();
		final int n = _sortedList.size();
		if (n == 0) { //nothing to pick
			return null;
		}
		final double min = _sortedList.get(0);
		final double max = _sortedList.get(_sortedList.size()-1);
		return Double.valueOf(min + rank0 / 100 * (max - min));
	}
	
	// see https://en.wikipedia.org/wiki/Percentile#Worked_Examples_of_the_Second_Variant
	// PERCENTILE(array, rank0/100) == PERCENTILE.INC(array, rank0/100)
	private Double _pickPercentileInc(double rank0) {
		_prepareSortedList();
		final int n = _sortedList.size();
		if (n == 0) { //nothing to pick
			return null;
		}
		final double x = rank0 / 100 * (n - 1) + 1; // percentile rank
		final int i = (int) Math.floor(x); // ith  
		double frac = x - i;
		double vi = _sortedList.get(i-1); // ith value
		double vi1 = _sortedList.get(i);  // (i+1)th value
		return Double.valueOf(vi + frac * (vi1 - vi)); // interpolated percentile value
	}
}
