package org.zkoss.zss.ngmodel.chart;

public interface NCategoryChartData extends NChartData{

	public int getNumOfSeries();
	public NSeries getSeriesAt(int i);
	
	public int getNumOfCategory();
	public NCategory getCategoryAt(int i);
}
