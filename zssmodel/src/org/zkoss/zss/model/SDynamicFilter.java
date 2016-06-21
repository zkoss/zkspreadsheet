/* SDynamicFilter.java

	Purpose:
		
	Description:
		
	History:
		May 16, 2016 6:25:07 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import org.zkoss.zss.model.SAutoFilter.FilterOp;

//ZSS-1226
/**
 * @author henri
 * @since 3.9.0
 */
public interface SDynamicFilter {
	Double getValue();
	Double getMaxValue();
	String getType(); //ZSS-1234
}
