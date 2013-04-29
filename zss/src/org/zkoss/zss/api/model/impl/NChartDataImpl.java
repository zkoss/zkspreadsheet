package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.zss.api.model.NChartData;

public class NChartDataImpl implements NChartData{

	ModelRef<ChartData> chartDataRef;
	
	public NChartDataImpl(ModelRef<ChartData> chartDataRef) {
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
		NChartDataImpl other = (NChartDataImpl) obj;
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
