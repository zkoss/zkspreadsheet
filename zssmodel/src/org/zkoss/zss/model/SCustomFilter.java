/* SCustomFilter.java

	Purpose:
		
	Description:
		
	History:
		May 11, 2016 11:38:42 AM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

//ZSS-1224
/**
 * @author henri
 * @since 3.9.0
 */
public interface SCustomFilter {
	String getValue();
	Operator getOperator();
	
	enum Operator {
		equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual, 
		beginWith, notBeginWith, endWith, notEndWith, contains, notContains,
	}
}
