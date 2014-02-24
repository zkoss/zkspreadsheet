/* ChartData.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model;

import org.zkoss.zss.ngmodel.chart.NChartData;

/**
 * This interface provides the access to the underlying data object ({@link NChartData}) of a chart.
 * @author dennis
 * @since 3.0.0
 */
public interface ChartData {

	/**
	 * Get the internal chart data
	 * @since 3.5.0
	 */
	public NChartData getInternalChartData();
}
