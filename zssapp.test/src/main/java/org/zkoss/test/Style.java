/* Style.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 20, 2012 9:19:15 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

/**
 * @author sam
 *
 */
public enum Style {
	PADDINGS,
	MARGINS,
	BORDERS;

	@Override
	public String toString() {
		return super.toString().toLowerCase();
	}
}
