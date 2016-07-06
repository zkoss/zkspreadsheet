/* CellMatch2.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 6:24:29 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SConditionalFormattingRule.RuleOperator;
import org.zkoss.zss.model.SCustomFilter;
import org.zkoss.zss.model.impl.RuleInfo;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.formula.FormulaEngine;

//ZSS-1142
/**
 * Use to check if a cell's value match the specified condition (used in ConditionalFormatting)
 * @author henri
 * @since 3.9.0
 */
public class CellMatch2 implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = 4775550316177451879L;
	private MatchOne f1;
	private MatchOne f2;
	private boolean isAnd;
	public CellMatch2(RuleOperator op1, String text1) {
		this.f1 = new MatchOne(op1, text1);
		this.f2 = null;
		this.isAnd = false;
	}
	
	public CellMatch2(RuleOperator op1, RuleInfo rule1, 
			RuleOperator op2, RuleInfo rule2, boolean isAnd) {
		this.f1 = new MatchOne(op1, rule1);
		this.f2 = op2 == null ? null : new MatchOne(op2, rule2);
		this.isAnd = isAnd;
	}
	
	public boolean match(SCell cell) {
		final boolean matchf1 = f1.match(cell);
		if (isAnd && !matchf1) {
			return false;
		}
		if (!isAnd && matchf1) {
			return true;
		}
		if (f2 != null) {
			return f2.match(cell); 
		} else {
			return matchf1;
		}
	}
	
	static class MatchOne implements Serializable {
		private static final long serialVersionUID = -7474902362367302742L;
		private boolean _not;
		private Matchable<SCell> _matchable;
		private RuleOperator op;
		private FormatEngine _formatEngine;
		
		MatchOne(RuleOperator op, String text) {
			this.op = op;
			_prepareMatch(null, text);
		}

		MatchOne(RuleOperator op, RuleInfo ruleInfo) {
			this.op = op;
			_prepareMatch(ruleInfo, null);
		}

		public String _getFormattedText(SCell cell){
			return getFormatEngine().format(cell, new FormatContext(ZssContext.getCurrent().getLocale())).getText();
		}
		
		protected FormatEngine getFormatEngine(){
			if(_formatEngine==null){
				_formatEngine = EngineFactory.getInstance().createFormatEngine();
			}
			return _formatEngine;
		}
		
		public boolean match(SCell cell) {
			final boolean ret = match0(cell);
			return _not ? !ret : ret;
		}
		
		private boolean match0(SCell cell) {
			if (_matchable != null) {
				return ((Matchable<SCell>)_matchable).match(cell);
			} else {
				return false;
			}
		}
		
		private void _prepareMatch(RuleInfo ruleInfo, String value) {
			switch(op) {
			case BEGINS_WITH:
				_matchable = ruleInfo == null ?
						new BeginsWith(value.toString()):
						new BeginsWith(ruleInfo);
				break;
			case ENDS_WITH:
				_matchable = ruleInfo == null ? 
						new EndsWith(value.toString()):
						new EndsWith(ruleInfo);
				break;

			case NOT_CONTAINS:
				_not = true;
			case CONTAINS_TEXT:
				_matchable = ruleInfo == null ? 
						new ContainsText(value.toString()):
						new ContainsText(ruleInfo);
				break;
				
			case NOT_EQUAL:
				_not = true;
			case EQUAL:
				_matchable = ruleInfo == null ?
						new Equals(value.toString()):
						new Equals(ruleInfo);
				break;
				
			case LESS_THAN_OR_EQUAL:
				_not = true;
			case GREATER_THAN:
				_matchable = ruleInfo == null ? 
						new GreaterThan2(value.toString()):
						new GreaterThan2(ruleInfo);
				break;

			case GREATER_THAN_OR_EQUAL:
				_not = true;
			case LESS_THAN:
				_matchable = ruleInfo == null ?
						new LessThan2(value.toString()):
						new LessThan2(ruleInfo);
				break;
			}
		}
	}
}
