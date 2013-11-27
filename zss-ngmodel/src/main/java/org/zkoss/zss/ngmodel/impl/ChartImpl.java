package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.chart.NChartData;

public class ChartImpl extends AbstractChart {
	private static final long serialVersionUID = 1L;
	String id;
	NChartType type;
	NViewAnchor anchor;
	NChartData data;
	String title;
	String xAxisTitle;
	String yAxisTitle;
	
	public ChartImpl(String id,NChartType type,NChartData data,NViewAnchor anchor){
		this.id = id;
		this.data = data;
		this.type = type;
		this.anchor = anchor;
	}
	
	public String getId() {
		return id;
	}
	public NViewAnchor getAnchor() {
		return anchor;
	}
	public void setAnchor(NViewAnchor anchor) {
		this.anchor = anchor;
	}
	public NChartType getType(){
		return type;
	}
	public NChartData getData() {
		return data;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getXAxisTitle() {
		return xAxisTitle;
	}

	public void setXAxisTitle(String xAxisTitle) {
		this.xAxisTitle = xAxisTitle;
	}

	public String getYAxisTitle() {
		return yAxisTitle;
	}

	public void setYAxisTitle(String yAxisTitle) {
		this.yAxisTitle = yAxisTitle;
	}

	
}
