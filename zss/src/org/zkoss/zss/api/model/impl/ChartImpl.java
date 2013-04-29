package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.model.sys.XBook;

public class ChartImpl implements Chart{
	
	ModelRef<XBook> bookRef;
	ModelRef<org.zkoss.poi.ss.usermodel.Chart> chartRef;
	
	public ChartImpl(ModelRef<XBook> bookRef, ModelRef<org.zkoss.poi.ss.usermodel.Chart> chartRef) {
		this.bookRef = bookRef;
		this.chartRef = chartRef;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((chartRef == null) ? 0 : chartRef.hashCode());
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
		ColorImpl other = (ColorImpl) obj;
		if (chartRef == null) {
			if (other.colorRef != null)
				return false;
		} else if (!chartRef.equals(other.colorRef))
			return false;
		return true;
	}
	
	public org.zkoss.poi.ss.usermodel.Chart getNative() {
		return chartRef.get();
	}
}
