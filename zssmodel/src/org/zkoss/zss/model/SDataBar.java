/* SDataBar.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 4:32:28 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;

/**
 * Condiditional Formatting of "dataBar" type.
 * @author henri
 * @since 3.8.2
 */
public interface SDataBar {
	/**
	 * Returns the value objects.
	 * @return
	 */
	List<SCFValueObject> getCFValueObjects();
	
	/** Returns the color of the bar */
	SColor getColor();
	
	/** Returns the minimum length of this data bar */
	Long getMinLength();
	
	/** Returns the maximum length of this data bar */
	Long getMaxLength();
	
	/** Returns whether show the value in the data bar */
	boolean isShowValue();
}
