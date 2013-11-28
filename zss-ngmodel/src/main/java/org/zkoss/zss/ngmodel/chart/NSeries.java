package org.zkoss.zss.ngmodel.chart;

public interface NSeries {

	public String getName();
	
	public int getNumOfValue();
	public Object getValueAt(int index);
	
	//for Scatter, xy chart
	public int getNumOfXValue();
	public Object getXValueAt(int index);
	public int getNumOfYValue();
	public Object getYValueAt(int index);
	
	
	public String getNameFormula();
	public void setNameFormula(String expr);
	public String getValuesFormula();
	public void setValuesFormula(String expr);
	public String getXValuesFormula();
	public void setXValuesFormula(String expr);
	public String getYValuesFormula();
	public void setYValuesFormula(String expr);
}
