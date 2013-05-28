package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.ChartData;

public class ChartDataImpl implements ChartData{

	ModelRef<org.zkoss.poi.ss.usermodel.charts.ChartData> chartDataRef;
	
	public ChartDataImpl(ModelRef<org.zkoss.poi.ss.usermodel.charts.ChartData> chartDataRef) {
		this.chartDataRef = chartDataRef;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result
				+ ((chartDataRef == null) ? 0 : chartDataRef.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		ChartDataImpl other = (ChartDataImpl) obj;
		if (chartDataRef == null) {
			if (other.chartDataRef != null)
				return false;
		} else if (!chartDataRef.equals(other.chartDataRef))
			return false;
		return true;
	}
	
	public org.zkoss.poi.ss.usermodel.charts.ChartData getNative(){
		return chartDataRef.get();
	}
	
	
	public org.zkoss.poi.ss.usermodel.charts.ChartData getPoiChartData(){
		return chartDataRef.get();
	}

}
