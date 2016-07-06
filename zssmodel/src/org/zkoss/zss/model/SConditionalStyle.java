/* SConditionalStyle.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 3:54:19 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

/**
 * Extra embedded style for SConditonalFormatting.
 * 
 * @author henri
 * @sicne 3.9.0
 */
public interface SConditionalStyle extends SExtraStyle {
	public SColorScale getColorScale();

	public SDataBar getDataBar();

	public Double getBarPercent();

	public SIconSet getIconSet();

	public Integer getIconSetId();
}
