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
		
		public Set getCriteria1();

		public Set getCriteria2();
				
		public boolean isShowButton(); //show filter button
		
		public FilterOp getOperator();
		
		public void setProperties(FilterOp filterOp, Object criteria1, Object criteria2, Boolean showButton);

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
	 * @param index the nth column (1st column in the filter range is 0)
	 * @return the column filter information of the specified column; null if the column is not filtered.
	 */
	public NFilterColumn getFilterColumn(int index, boolean create);
	
	public void clearFilterColumn(int index);
	
	public void clearFilterColumns();
}
