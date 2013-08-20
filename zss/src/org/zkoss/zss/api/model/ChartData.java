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

/**
 * This interface provides the access to the underlying data object ({@link org.zkoss.poi.ss.usermodel.charts.ChartData}) of a chart.
 * @author dennis
 * @since 3.0.0
 */
public interface ChartData {

	public org.zkoss.poi.ss.usermodel.charts.ChartData getPoiChartData();
}
