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
 * @author Sam, Ian Tsai
 *
 */
public interface CellFilter {
	/**
	 * 
	 * @param context
	 * @return
	 */
	public boolean doFilter(CellVisitorContext context);
}
