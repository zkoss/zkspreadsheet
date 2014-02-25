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
package org.zkoss.zss.model;

import org.zkoss.zss.model.chart.SChartData;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface SChart {

	public enum ChartType{
		AREA,
		BAR,
		BUBBLE,
		COLUMN,
		DOUGHNUT,
		LINE,
		OF_PIE,
		PIE,
		RADAR,
		SCATTER,
		STOCK,
		SURFACE
	}
	
	public enum ChartGrouping {
		STANDARD,
		STACKED,
		PERCENT_STACKED,
		CLUSTERED; //bar only
	}
	
	public enum ChartLegendPosition {
		BOTTOM,
		LEFT,
		RIGHT,
		TOP,
		TOP_RIGHT
	}
	public enum BarDirection {
		HORIZONTAL, //horizontal, bar chart
		VERTICAL; //vertical, column chart
	}

	
	public String getId();
	
	public SSheet getSheet();
	
	public ChartType getType();
	
	public SChartData getData();
	
	public ViewAnchor getAnchor();
	
	public void setAnchor(ViewAnchor anchor);
	
	public String getTitle();
	
	public String getXAxisTitle();
	
	public String getYAxisTitle();

	void setTitle(String title);

	void setXAxisTitle(String xAxisTitle);

	void setYAxisTitle(String yAxisTitle);
	
	public void setLegendPosition(ChartLegendPosition pos);
	
	public ChartLegendPosition getLegendPosition();
	
	public void setGrouping(ChartGrouping grouping);
	
	public ChartGrouping getGrouping();
	
	public BarDirection getBarDirection();
	public void setBarDirection(BarDirection direction);
	
	public boolean isThreeD();
	
	public void setThreeD(boolean threeD);
}
