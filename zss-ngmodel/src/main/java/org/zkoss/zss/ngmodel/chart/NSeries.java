package org.zkoss.zss.ngmodel.chart;

public interface NSeries {

	public String getName();
	
	public int getNumOfValue();
	public Object getValue(int index);
	
	//for Scatter, xy chart
	public int getNumOfYValue();
	public Object getYValue(int index);
	
	public void setFormula(String nameExpr,String valuesExpr,String yValuesExpr);
	public String getNameFormula();
	public String getValuesFormula();
	public String getYValuesFormula();
	
	public void clearFormulaResultCache();
}
