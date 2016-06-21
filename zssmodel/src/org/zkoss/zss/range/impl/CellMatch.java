/* CellMatch.java

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
import org.zkoss.zss.model.SCustomFilter;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.formula.FormulaEngine;

//ZSS-1192
/**
 * Use to check if a cell's value match the specified CustomFilter
 * @author henri
 * @since 3.9.0
 */
public class CellMatch implements Matchable<SCell>, Serializable {
	private static final long serialVersionUID = 4775550316177451879L;
	
	private FilterMatch f1;
	private FilterMatch f2;
	private boolean isAnd;
	public CellMatch(SCustomFilter f1, SCustomFilter f2, boolean isAnd) {
		this.f1 = new FilterMatch(f1);
		this.f2 = f2 == null ? null : new FilterMatch(f2);
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
	
	static class FilterMatch implements Serializable {
		private boolean _not;
		private Pattern _pattern;
		private Matchable<?> _matchable;
		private Object value;
		private FilterOp op;
		private CellType type;
		private FormatEngine _formatEngine;
		
		FilterMatch(SCustomFilter filter) {
			op = filter.getOperator();
			// Match depends on the Operator
			// equal/notEqual/beginWidth/notBeginWidth/endWith/notEndWidth/contains/notContains: always check as a String
			// greaterThan/greaterThanOrEqual/lessThan/lessThanOrEqual: depends on the input type
			switch(op) {
			case equal:
			case notEqual:
			case beginWith:
			case notBeginWith:
			case endWith:
			case notEndWith:
			case contains:
			case notContains:
				value = filter.getValue();
				type = CellType.STRING;
				break;
				
			default:
				try {
					value = Double.valueOf(filter.getValue());
					type = CellType.NUMBER;
				} catch (NumberFormatException ex) {
					value = filter.getValue();
					type = CellType.STRING;
				}
			}
			
			_prepareMatch();
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
			final CellType t = 
					(cell == null || cell.isNull() ? CellType.BLANK :
						cell.getType() == CellType.FORMULA ? 
						cell.getFormulaResultType() : cell.getType());
						
			if (t != type 
					&& op != FilterOp.equal 
					&& op != FilterOp.notEqual) {
				return _not;
			}
			final boolean ret = type == CellType.STRING ? 
					matchString(cell) : matchDouble(cell);
			return _not ? !ret : ret;
		}
		
		private boolean matchString(SCell cell) {
			final String formattedText = _getFormattedText(cell);
			if (_pattern != null) {
				final Matcher matcher = _pattern.matcher(formattedText);
				return matcher.matches();
			} else if (_matchable != null) {
				return ((Matchable<String>)_matchable).match(formattedText);
			} else {
				return false;
			}
		}
		
		private boolean matchDouble(SCell cell) {
			final Double val = cell.getNumberValue();
			return ((Matchable<Double>)_matchable).match(val);
		}
		
		private void _prepareMatch() {
			switch(op) {
			case notBeginWith:
				_not = true;
			case beginWith:
			{
				String val = _escape(value.toString());
				val = val + ".*";
				_pattern = Pattern.compile(val);
				break;
			}	
			case notEndWith:
				_not = true;
			case endWith:
			{
				String val = _escape(value.toString());
				val = ".*" + val;
				_pattern = Pattern.compile(val);
				break;
			}				
			case notContains:
				_not = true;
			case contains:
			{
				String val = _escape(value.toString());
				val = ".*" + val + ".*";
				_pattern = Pattern.compile(val);
				break;
			}				
			case notEqual:
				_not = true;
			case equal:
			{
				String val = _escape(value.toString());
				_pattern = Pattern.compile(val);
				break;
			}
			case after: //ZSS-1234
			case greaterThan:
				_matchable = type == CellType.STRING ? 
						new GreaterThan<String>(value.toString()):
						new GreaterThan<Double>((Double) value);
				break;
				
			case afterEq: //ZSS-1234
			case greaterThanOrEqual:
				_matchable = type == CellType.STRING ?
						new GreaterThanOrEqual<String>(value.toString()):
						new GreaterThanOrEqual<Double>((Double) value);
				break;
				
			case before: //ZSS-1234
			case lessThan:
				_matchable = type == CellType.STRING ?
						new LessThan<String>(value.toString()):
						new LessThan<Double>((Double) value);
				break;

			case beforeEq: //ZSS-1234
			case lessThanOrEqual:
				_matchable = type == CellType.STRING ?
						new LessThanOrEqual<String>(value.toString()):
						new LessThanOrEqual<Double>((Double) value);
				break;
			}
		}
		
		private String _escape(String src) {
			String s1 = src.replaceAll("([.+^$\\[\\]\\\\(){}|-])", "\\\\$0");
			String s2 = s1.replaceAll("\\*", ".*");
			return s2.replaceAll("\\?", ".");
		}
	}
	
}
