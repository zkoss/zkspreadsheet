package org.zkoss.zss.ngmodel.chart;

public interface NCategoryChartData extends NChartData{

	public int getNumOfSeries();
	public NSeries getSeriesAt(int i);
	
	public int getNumOfCategory();
	public Object getCategoryAt(int i);
	
	
	
	public NSeries addSeries();
	public void removeSeries(NSeries series);
	
	public void setCategoriesExpression(String expr);
}
