/* CellMatch.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 6:24:29 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl.autofill;

import java.io.Serializable;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.zkoss.poi.ss.usermodel.ZssContext;
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
public class StringCellMatch implements Serializable {
	private FilterMatch f1;
	private FilterMatch f2;
	private boolean isAnd;
	public StringCellMatch(SCustomFilter f1, SCustomFilter f2, boolean isAnd) {
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
		private Matchable<String> _matchable;
		private String value;
		private SCustomFilter.Operator op;
		private CellType type;
		private FormatEngine _formatEngine;
		
		FilterMatch(SCustomFilter filter) {
			value = filter.getValue();
			type = CellType.STRING;
			op = filter.getOperator();
			
			_prepareStringMatch();
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
					&& op != SCustomFilter.Operator.equal 
					&& op != SCustomFilter.Operator.notEqual) {
				return _not;
			}
			final boolean ret = match0(cell);
			return _not ? !ret : ret;
		}
		
		private boolean match0(SCell cell) {
			final String formattedText = _getFormattedText(cell);
			if (_pattern != null) {
				final Matcher matcher = _pattern.matcher(formattedText);
				return matcher.matches();
			} else if (_matchable != null) {
				return _matchable.match(formattedText);
			} else {
				return false;
			}
		}
		
		private void _prepareStringMatch() {
			switch(op) {
			case notBeginWith:
				_not = true;
			case beginWith:
			{
				String val = _escape(value);
				val = val + ".*";
				_pattern = Pattern.compile(val);
				break;
			}	
			case notEndWith:
				_not = true;
			case endWith:
			{
				String val = _escape(value);
				val = ".*" + val;
				_pattern = Pattern.compile(val);
				break;
			}				
			case notContains:
				_not = true;
			case contains:
			{
				String val = _escape(value);
				val = ".*" + val + ".*";
				_pattern = Pattern.compile(val);
				break;
			}				
			case notEqual:
				_not = true;
			case equal:
			{
				String val = _escape(value);
				_pattern = Pattern.compile(val);
				break;
			}				
			case greaterThan:
				_matchable = new GreaterThan<String>(value);
				break;
				
			case greaterThanOrEqual:
				_matchable = new GreaterThanOrEqual<String>(value);
				break;
				
			case lessThan:
				_matchable = new LessThan<String>(value);
				break;

			case lessThanOrEqual:
				_matchable = new LessThanOrEqual<String>(value);
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

final class GreaterThan<T extends Comparable<T>> implements Matchable<T> {
	final private T base;
	GreaterThan(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) > 0;
	}
}

final class GreaterThanOrEqual<T extends Comparable<T>> implements Matchable<T> {
	final private T base;
	GreaterThanOrEqual(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) >= 0;
	}
}

final class LessThan<T extends Comparable<T>> implements Matchable<T> {
	final private T base;
	LessThan(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) < 0;
	}
}

final class LessThanOrEqual<T extends Comparable<T>> implements Matchable<T> {
	final private T base;
	LessThanOrEqual(T b) {
		base = b;
	}
	@Override
	public boolean match(T value) {
		return value.compareTo(base) <= 0;
	}
}
