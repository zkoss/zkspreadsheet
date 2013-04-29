package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.zss.api.model.NChart;
import org.zkoss.zss.model.sys.Book;

public class NChartImpl implements NChart{
	
	ModelRef<Book> bookRef;
	ModelRef<Chart> chartRef;
	
	public NChartImpl(ModelRef<Book> bookRef, ModelRef<Chart> chartRef) {
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
		NColorImpl other = (NColorImpl) obj;
		if (chartRef == null) {
			if (other.colorRef != null)
				return false;
		} else if (!chartRef.equals(other.colorRef))
			return false;
		return true;
	}
	
	public Chart getNative() {
		return chartRef.get();
	}
}
