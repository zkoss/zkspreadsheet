/* RuleInfo.java

	Purpose:
		
	Description:
		
	History:
		Jun 24, 2016 4:38:14 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.ss.formula.LazyAreaEval;
import org.zkoss.poi.ss.formula.LazyRefEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.ptg.AreaPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.ptg.RefPtg;
import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;

/**
 * @author Henri
 *
 */
public class RuleInfo {
	private SSheet _sheet;
	private ConditionalFormattingRuleImpl _rule;
	private int row0;
	private int col0;
	private FormulaExpression _formulaExpr;

	private Map<String, EvaluationResult> _cacheMap; 
	
	RuleInfo(SSheet sheet, ConditionalFormattingRuleImpl rule, FormulaExpression formulaExpr, int row0, int col0) {
		this._sheet = sheet;
		this._rule = rule;
		this._formulaExpr = formulaExpr;
		this.row0 = row0;
		this.col0 = col0;
		_cacheMap = new HashMap<String, EvaluationResult>();
		if (formulaExpr != null) {
			rule.addDependency(formulaExpr);
		}
	}
	
	public EvaluationResult evalFormula(SCell cell){
		synchronized (this) {
			final int row = cell == null ? row0 : cell.getRowIndex();
			final int col = cell == null ? col0 : cell.getColumnIndex();
			final String key = ""+row+"_"+col;
			final EvaluationResult cache = _cacheMap.get(key);
			if (cache != null) {
				return cache;
			}
			
			final int rowOffset = row - row0;
			final int colOffset = col - col0;
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			if(_formulaExpr!=null) {
//				if (!_rule.hasNoRelative()) {
//					if (prepareDependency(rowOffset, colOffset)) {
//						_rule.setNoRelative(true);
//					}
//				}
				final EvaluationResult result = 
					fe.evaluate(_formulaExpr,new FormulaEvaluationContext(
							_sheet, cell, null, new int[] {rowOffset, colOffset})); //ZSS-1257
				_cacheMap.put(key, result);
				return result;
			}
			return null;
		}
	}
	
//	public void clearCache(int row, int col) {
//		_cacheMap.remove(""+row+"_"+col);
//	}
	
	public void clearCacheMap() {
		_cacheMap = new HashMap<String, EvaluationResult>();
	}
}
