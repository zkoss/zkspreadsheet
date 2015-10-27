/* SConditionalFormatting.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 11:51:24 AM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;

/**
 * Conditional Formatting
 * @author henri
 *
 */
public interface SConditionalFormatting {
	/**
	 * The sheet on which this conditional formatting covered
	 * @return
	 */
	SSheet getSheet();
	
	/**
	 * Regions that this conditional formatting covered
	 * @return
	 */
	List<CellRegion> getRegions();
	
	/**
	 * Rules applied to the covered region
	 * @return
	 */
	List<SConditionalFormattingRule> getRules();
}
