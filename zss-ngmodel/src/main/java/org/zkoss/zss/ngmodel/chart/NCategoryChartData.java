package org.zkoss.zss.ngmodel.chart;

public interface NCategoryChartData extends NChartData{

	/**
	 * Return formula parsing state.
	 * @return true if has error, false if no error or no formula
	 */
	public boolean isFormulaParsingError();
	
	public int getNumOfSeries();
	public NSeries getSeries(int i);
	
	public int getNumOfCategory();
	public Object getCategory(int i);
	
	
	
	public NSeries addSeries();
	public void removeSeries(NSeries series);
	
	public String getCategoriesFormula();
	public void setCategoriesFormula(String expr);
	
}
