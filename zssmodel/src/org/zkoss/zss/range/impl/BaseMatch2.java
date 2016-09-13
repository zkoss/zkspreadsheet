/* GreaterThan.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:13:53 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;
import java.util.Date;

import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SRichText;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.impl.CellValue;
import org.zkoss.zss.model.impl.RuleInfo;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;

/**
 * @author henri
 * @since 3.9.0
 */
abstract public class BaseMatch2 implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = -2673045590406268437L;
	final private Object base;
	final private RuleInfo ruleInfo;
	
	public BaseMatch2(Object b) {
		base = b;
		ruleInfo = null;
	}
	public BaseMatch2(RuleInfo ruleInfo) {
		base = null;
		this.ruleInfo = ruleInfo;
	}
		
	@Override
	public boolean match(SCell cell) {
//ZSS-1270: mark out; cell could be blank
//		if (cell == null || cell.isNull()) {
//			return false;
//		}
		if (base != null) {
			if (base instanceof String) {
				return matchString0(cell, (String)base);
			} else if (base instanceof Number) {
				return matchDouble0(cell, (Double)base);
			} 
		} else {
			final EvaluationResult result = ruleInfo.evalFormula(cell);
			if(result.getType() == ResultType.SUCCESS){
				final Object val = result.getValue();
				if (val instanceof Double) {
					return matchDouble0(cell, (Double)val);
				}
				if (val instanceof Date) {
					return matchDouble0(cell, Double.valueOf(DateUtil.getExcelDate((Date)val)));
				}
				if (val instanceof String) {
					return matchString0(cell, (String)val);
				}
			}
			//blank, error, boolean is deemed as false
		}
		return false;
	}
	
	private boolean matchString0(SCell cell, String b) {
		Object value =  CellValueHelper.inst.getValue(cell);
		//ZSS-1270
		if (value == null) {
			value = "";
		}
		if (value instanceof String) {
			final String formattedText = CellValueHelper.inst.getFormattedText(cell);
			return matchString(formattedText, b);
		} else if (value instanceof SRichText) {
			final String text = ((SRichText)value).getText();
			return matchString(text, b);
		}
		return false;
	}
	
	private boolean matchDouble0(SCell cell, Double b) {
		final Object value =  CellValueHelper.inst.getValue(cell);
		if (value instanceof Double) {
			return matchDouble((Double)value, b);
		}
		return false;
	}
	
	abstract protected boolean matchString(String text, String b);
	
	abstract protected boolean matchDouble(Double num, Double b);
}
