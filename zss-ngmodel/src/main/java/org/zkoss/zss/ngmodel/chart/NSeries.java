package org.zkoss.zss.ngmodel.chart;

public interface NSeries {

	public String getName();
	public Object getValueAt(int index);
	
	//for Scatter
	public Object getXValueAt(int index);
	public Object getYValueAt(int index);
}
