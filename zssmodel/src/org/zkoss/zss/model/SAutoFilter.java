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
package org.zkoss.zss.model;

import java.util.Collection;
import java.util.List;
import java.util.Set;
import java.util.Map;

/**
 * Contains autofilter's setting.
 * @author Dennis
 * @since 3.5.0
 */
public interface SAutoFilter {

	/**
	 * A filter column contains information for filtering, e.g. criteria.
	 * A filter column only exists when users apply a criteria on a column. 
	 * @author Dennis
	 * @since 3.5.0
	 */
	public interface NFilterColumn{
		/**
		 * @return the nth column (1st column in the filter range is 0)
		 */
		int getIndex();
		
		public List<String> getFilters();
		
		/**
		 * @return main criteria used on a column
		 */
		public Set getCriteria1();

		public Set getCriteria2();
				
		public boolean isShowButton(); //show filter button
		
		public FilterOp getOperator();
		
		@Deprecated
		public void setProperties(FilterOp filterOp, Object criteria1, Object criteria2, Boolean showButton);

		//ZSS-1191
		//@since 3.9.0
		public void setProperties(FilterOp filterOp, Object criteria1, Object criteria2, Boolean showButton, Map<String, Object> extra);
		
		//ZSS-1191
		//@since 3.9.0
		public SColorFilter getColorFilter();
		
		//ZSS-1224
		//@since 3.9.0
		public SCustomFilters getCustomFilters();
		
		//ZSS-1226
		//@since 3.9.0
		public SDynamicFilter getDynamicFilter();
		
		//ZSS-1227
		//@since 3.9.0
		public STop10Filter getTop10Filter();
		
		//ZSS-1229
		//@since 3.9.0
		public boolean isFiltered();
	}
	
	/**
	 * 
	 * @author Dennis
	 * @since 3.5.0
	 */
	public enum FilterOp{
		equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual, 
		beginWith, notBeginWith, endWith, notEndWith, contains, notContains, none, custom,
		between, before, after, beforeEq, afterEq, betweenDates,

		and, bottom, or, top10, values,
		//ZSS-1191
		cellColor, fontColor,
		//ZSS-1226
		aboveAverage, belowAverage,
		//ZSS-1234
		tomorrow, today, yesterday, nextWeek, thisWeek, lastWeek,
		nextMonth, thisMonth, lastMonth, nextQuarter, thisQuarter, lastQuarter,
		nextYear, thisYear, lastYear, yearToDate, within,
		Q1, Q2, Q3, Q4,	M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11, M12, allDatesInPeriod,
		
	}
	

	/**
	 * Returns the filtered Region.
	 */
	public CellRegion getRegion();
	
	/**
	 * Return filter setting of each filtered column.
	 */
	public Collection<NFilterColumn> getFilterColumns();
	
	/**
	 * Returns the column filter information of the specified column; null if the column is not filtered.
	 * @param index the nth column (1st column in the filter range is 0)
	 * @return the column filter information of the specified column; null if the column is not filtered.
	 */
	public NFilterColumn getFilterColumn(int index, boolean create);
	
	public void clearFilterColumn(int index);
	
	public void clearFilterColumns();
	
	//ZSS-1229
	//@since 3.9.0
	public boolean isFiltered();
}
