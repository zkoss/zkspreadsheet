/* SortSelector.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 13, 2010 6:35:28 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.sort;

import org.zkoss.util.resource.Labels;

/**
 * @author Sam
 *
 */
public final class SortSelector {
	private SortSelector(){}
	
	
	/**
	 * Returns sort order base on i3-label
	 * <p> true to do descending sort; false to do ascending sort
	 * <p> Default: false
	 * @param label
	 * @return
	 */
	public static boolean getSortOrder(String i3label) {
		if (i3label == null)
			return false;
		
		if ( i3label.equals(Labels.getLabel("sort.ascending"))
				|| i3label.equals(Labels.getLabel("sort.str.ascending"))
				|| i3label.equals(Labels.getLabel("sort.num.ascending")))
			return true;
		return false;
	}
}
