/* CellFilter.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 17, 2010 5:15:14 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

/**
 * Internal Use Only.
 * @author Sam, Ian Tsai
 *
 */
public interface CellFilter {
	/**
	 * Returns whether filter this cell.
	 * @param context
	 * @return whether filter this cell
	 */
	public boolean doFilter(CellVisitorContext context);
}
