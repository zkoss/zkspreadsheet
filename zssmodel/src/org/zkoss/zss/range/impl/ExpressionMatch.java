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
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.impl.CellValue;
import org.zkoss.zss.model.impl.RuleInfo;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;

/**
 * @author henri
 * @since 3.9.0
 */
public class ExpressionMatch implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = -1931445210569317355L;
	final private RuleInfo ruleInfo;
	
	public ExpressionMatch(RuleInfo ruleInfo) {
		this.ruleInfo = ruleInfo;
	}
	
	@Override
	public boolean match(SCell cell) {
		final EvaluationResult result = ruleInfo.evalFormula(cell);
		if(result.getType() == ResultType.SUCCESS){
			final Object val = result.getValue();
			if (val instanceof Double) {
				return ((Double)val).doubleValue() != 0;
			}
			if (val instanceof Boolean) {
				return ((Boolean)val).booleanValue();
			}
			if (val instanceof Date) {
				return DateUtil.getExcelDate((Date)val) != 0;
			}
		}
		//blank, string and error is false
		return false;
	}
}
