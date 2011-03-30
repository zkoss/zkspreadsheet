/* FullWeekData.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 5:21:07 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

/**
 * @author henrichen
 *
 */
public class FullWeekData extends CircularData {
	private static final String[] WEEKKEYS = new String[] {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
	public FullWeekData(int type) {
		super(WEEKKEYS, type);
	}
}
