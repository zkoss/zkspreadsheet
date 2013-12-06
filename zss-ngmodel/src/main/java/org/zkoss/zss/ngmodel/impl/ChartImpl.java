/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.chart.NChartData;
import org.zkoss.zss.ngmodel.impl.chart.CategoryChartDataImpl;
import org.zkoss.zss.ngmodel.impl.chart.ChartDataAdv;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ChartImpl extends ChartAdv {
	private static final long serialVersionUID = 1L;
	String id;
	NChartType type;
	NViewAnchor anchor;
	ChartDataAdv data;
	String title;
	String xAxisTitle;
	String yAxisTitle;
	SheetAdv sheet;
	
	public ChartImpl(SheetAdv sheet,String id,NChartType type,NViewAnchor anchor){
		this.sheet = sheet;
		this.id = id;
		this.type = type;
		this.anchor = anchor;
		this.data = createChartData(type);
	}
	@Override
	public NSheet getSheet(){
		checkOrphan();
		return sheet;
	}
	@Override
	public String getId() {
		return id;
	}
	@Override
	public NViewAnchor getAnchor() {
		return anchor;
	}
	@Override
	public NChartType getType(){
		return type;
	}
	@Override
	public NChartData getData() {
		return data;
	}
	@Override
	public String getTitle() {
		return title;
	}
	@Override
	public void setTitle(String title) {
		this.title = title;
	}
	@Override
	public String getXAxisTitle() {
		return xAxisTitle;
	}
	@Override
	public void setXAxisTitle(String xAxisTitle) {
		this.xAxisTitle = xAxisTitle;
	}
	@Override
	public String getYAxisTitle() {
		return yAxisTitle;
	}
	@Override
	public void setYAxisTitle(String yAxisTitle) {
		this.yAxisTitle = yAxisTitle;
	}

	private ChartDataAdv createChartData(NChartType type){
		//TODO for type;
		switch(type){
		case AREA:
		case BAR:
		case COLUMN:
		case LINE:
		case PIE:
			return new CategoryChartDataImpl(this,id+"Data");
		}
		throw new UnsupportedOperationException("unsupported chart type "+type);
	}

	@Override
	public void destroy() {
		checkOrphan();
		
		((ChartDataAdv)getData()).destroy();
		
		sheet = null;
	}
	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	
}
