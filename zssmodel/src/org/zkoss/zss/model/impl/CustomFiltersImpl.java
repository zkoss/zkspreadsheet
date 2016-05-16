/* CustomFiltersImpl.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 11:47:43 AM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.model.SCustomFilter;
import org.zkoss.zss.model.SCustomFilters;

//ZSS-1224
/**
 * @author henri
 * @since 3.9.0
 */
public class CustomFiltersImpl implements SCustomFilters, Serializable {
	private int type = 2; //Date 0, Number 1, String 2
	private boolean and;
	private List<SCustomFilter> filters;
	
	public CustomFiltersImpl(SCustomFilter filter1, SCustomFilter filter2, boolean and) {
		this.and = and;
		filters = new ArrayList<SCustomFilter>(2);
		filters.add(filter1);
		if (filter2 != null) {
			filters.add(filter2);
		}
	}
	
	@Override
	public boolean isAnd() {
		return and;
	}

	@Override
	public SCustomFilter getCustomFilter1() {
		return filters.get(0);
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.model.SCustomFilters#getCustomFilter2()
	 */
	@Override
	public SCustomFilter getCustomFilter2() {
		return filters.size() > 1 ? filters.get(1) : null;
	}
	
	//ZSS-1183, ZSS-1224
	//@since 3.9.0
	/*package*/ CustomFiltersImpl cloneCustomFilters() {
		final CustomFilterImpl filter1 = (CustomFilterImpl)getCustomFilter1();
		final CustomFilterImpl filter2 = (CustomFilterImpl)getCustomFilter2();
		return new CustomFiltersImpl(filter1.cloneCustomFilter(), 
				filter2 == null ? null : filter2.cloneCustomFilter(), isAnd());
	}
}
