package org.zkoss.zss.ngmodel.chart;

public interface NCategoryChartData extends NChartData{

	public int getNumOfSeries();
	public NSeries getSeries(int i);
	
	public int getNumOfCategory();
	public Object getCategory(int i);
	
	
	
	public NSeries addSeries();
	public void removeSeries(NSeries series);
	
	public String getCategoriesFormula();
	public void setCategoriesFormula(String expr);
	
}
