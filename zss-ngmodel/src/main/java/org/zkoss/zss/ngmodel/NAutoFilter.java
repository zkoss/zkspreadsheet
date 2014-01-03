package org.zkoss.zss.ngmodel;

import java.util.Collection;
import java.util.List;
import java.util.Set;

public interface NAutoFilter {

	public interface NFilterColumn{
		/**
		 * the column index in the auto filter
		 * @return
		 */
		int getIndex();
		
		public List<String> getFilters();
		
		public void addFilter(String filter);
		
		public void clearFilter();
		
		public Set getCriteria1();
		
		public void addCriteria1(Object obj);
		
		public void clearCriteria1();
		
		public Set getCriteria2();
		
		public void addCriteria2(Object obj);
		
		public void clearCriteria2();		
		
		public boolean isOn();
		
		public void setOn(boolean on);
		
		public FilterOp getOperator();
		
		public void setOperator(FilterOp op);
	}
	
	public enum FilterOp{
		AND, BOTTOM10, BOTOOM10_PERCENT, OR, TOP10, TOP10_PERCENT, VALUES;
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
	 * @param col the nth column (1st column in the filter range is 0)
	 * @return the column filter information of the specified column; null if the column is not filtered.
	 */
	public NFilterColumn getFilterColumn(int index, boolean create);
	
	public void clearFilterColumn(int index);
	
	public void clearFilterColumns();
}
