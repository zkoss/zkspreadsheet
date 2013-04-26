package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.zss.model.Book;

public class NChart {

	public enum Type {
		Area3D,
		Area,
		Bar3D,
		Bar,
		Bubble,
		Column,
		Column3D,
		Doughnut,
		Line3D,
		Line,
		OfPie,
		Pie3D,
		Pie,
		Radar,
		Scatter,
		Stock,
		Surface3D,
		Surface,
	}
	public enum Grouping {
		STANDARD,
		STACKED,
		PERCENT_STACKED,
		CLUSTERED; //bar only
	}
	public enum LegendPosition {
		BOTTOM,
		LEFT,
		RIGHT,
		TOP,
		TOP_RIGHT
	}

	
	ModelRef<Book> bookRef;
	ModelRef<Chart> chartRef;
	
	public NChart(ModelRef<Book> bookRef, SimpleRef<Chart> chartRef) {
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
		NColor other = (NColor) obj;
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
