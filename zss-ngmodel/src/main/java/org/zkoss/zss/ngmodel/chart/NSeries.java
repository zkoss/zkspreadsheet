package org.zkoss.zss.ngmodel.chart;

public interface NSeries {

	public String getName();
	public Object getValueAt(int index);
	
	//for Scatter, xy chart
	public Object getXValueAt(int index);
	public Object getYValueAt(int index);
	
	
	//
	public void setNameExpression(String expr);
	public void setValuesExpresion(String expr);
	public void setXValuesExpression(String expr);
	public void setYValuesExpression(String expr);
}
