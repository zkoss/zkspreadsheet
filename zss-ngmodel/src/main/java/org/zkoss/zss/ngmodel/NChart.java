/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import org.zkoss.zss.ngmodel.chart.NChartData;
import org.zkoss.zss.ngmodel.impl.SheetAdv;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NChart {

	public enum NChartType{
		COLUMN,
		LINE,
		PIE,
		BAR,
		AREA
	}
	
	public String getId();
	
	public NSheet getSheet();
	
	public NChartType getType();
	
	public NChartData getData();
	
	public NViewAnchor getAnchor();
	
	public String getTitle();
	
	public String getXAxisTitle();
	
	public String getYAxisTitle();

	void setTitle(String title);

	void setXAxisTitle(String xAxisTitle);

	void setYAxisTitle(String yAxisTitle);
}
