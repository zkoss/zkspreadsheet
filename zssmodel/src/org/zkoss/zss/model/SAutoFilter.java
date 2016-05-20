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
	}
	
	/**
	 * 
	 * @author Dennis
	 * @since 3.5.0
	 */
	public enum FilterOp{
		AND, BOTTOM10, BOTTOM10_PERCENT, OR, TOP10, TOP10_PERCENT, VALUES, 
		CELL_COLOR, FONT_COLOR, //ZSS-1191
		ABOVE_AVERAGE, BELOW_AVERAGE; //ZSS-1226
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
}
