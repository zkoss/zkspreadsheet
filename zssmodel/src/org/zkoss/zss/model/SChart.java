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

	public enum NChartType{
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
	
	public enum NChartGrouping {
		STANDARD,
		STACKED,
		PERCENT_STACKED,
		CLUSTERED; //bar only
	}
	
	public enum NChartLegendPosition {
		BOTTOM,
		LEFT,
		RIGHT,
		TOP,
		TOP_RIGHT
	}
	public enum NBarDirection {
		HORIZONTAL, //horizontal, bar chart
		VERTICAL; //vertical, column chart
	}

	
	public String getId();
	
	public SSheet getSheet();
	
	public NChartType getType();
	
	public SChartData getData();
	
	public ViewAnchor getAnchor();
	
	public void setAnchor(ViewAnchor anchor);
	
	public String getTitle();
	
	public String getXAxisTitle();
	
	public String getYAxisTitle();

	void setTitle(String title);

	void setXAxisTitle(String xAxisTitle);

	void setYAxisTitle(String yAxisTitle);
	
	public void setLegendPosition(NChartLegendPosition pos);
	
	public NChartLegendPosition getLegendPosition();
	
	public void setGrouping(NChartGrouping grouping);
	
	public NChartGrouping getGrouping();
	
	public NBarDirection getBarDirection();
	public void setBarDirection(NBarDirection direction);
	
	public boolean isThreeD();
	
	public void setThreeD(boolean threeD);
}
