package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.charts.ChartData;

public class NChartData {

	ModelRef<ChartData> chartDataRef;
	
	public NChartData(ModelRef<ChartData> chartDataRef) {
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
		NChartData other = (NChartData) obj;
		if (chartDataRef == null) {
			if (other.chartDataRef != null)
				return false;
		} else if (!chartDataRef.equals(other.chartDataRef))
			return false;
		return true;
	}
	
	public ChartData getNative(){
		return chartDataRef.get();
	}

}
