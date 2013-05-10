package org.zkoss.zss.api.model;


public interface Chart {

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
}
