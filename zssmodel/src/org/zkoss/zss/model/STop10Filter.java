/* STop10Filter.java

	Purpose:
		
	Description:
		
	History:
		May 17, 2016 3:02:07 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

//ZSS-1227
/**
 * @author henri
 * @since 3.9.0
 */
public interface STop10Filter {
	boolean isTop(); // whether pick top10 or bottom10
	boolean isPercent(); //whether pick item or percent
	double getValue(); // number(count) of items
	double getFilterValue(); // cell content value(inclusive) that filtering
}
