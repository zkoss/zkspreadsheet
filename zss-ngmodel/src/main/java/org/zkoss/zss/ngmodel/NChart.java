package org.zkoss.zss.ngmodel;

import org.zkoss.zss.ngmodel.chart.NChartData;

public interface NChart {

	public enum NChartType{
		COLUMN,
		LINE,
		PIE,
		BAR,
		AREA
	}
	
	public String getId();
	
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
