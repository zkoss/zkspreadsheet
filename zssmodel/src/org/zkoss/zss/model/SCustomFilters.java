/* SCustomFilters.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 11:45:10 AM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

//ZSS-1224
/**
 * @author henri
 * @since 3.9.0
 */
public interface SCustomFilters {
	boolean isAnd();
	
	SCustomFilter getCustomFilter1();
	
	SCustomFilter getCustomFilter2();
}
