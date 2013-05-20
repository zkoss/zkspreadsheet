package org.zkoss.zss.api.model;


public interface Chart {

	public enum Type {
		AREA_3D,
		AREA,
		BAR_3D,
		BAR,
		BUBBLE,
		COLUMN,
		COLUMN_3D,
		DOUGHNUT,
		LINE_3D,
		LINE,
		OF_PIE,
		PIE_3D,
		PIE,
		RADAR,
		SCATTER,
		STOCK,
		SURFACE_3D,
		SURFACE
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
	
	public String getId();
}
