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
public class FullMonthData extends CircularData {
	private static final String[] MONTHKEYS = new String[] {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
	public FullMonthData(int type) {
		super(MONTHKEYS, type);
	}
}
