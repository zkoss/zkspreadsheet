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
import org.zkoss.zss.ngmodel.impl.chart.GeneralChartDataImpl;
import org.zkoss.zss.ngmodel.impl.chart.ChartDataAdv;
import org.zkoss.zss.ngmodel.impl.chart.UnsupportedChartDataImpl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ChartImpl extends AbstractChartAdv {
	private static final long serialVersionUID = 1L;
	String id;
	NChartType type;
	NViewAnchor anchor;
	ChartDataAdv data;
	String title;
	String xAxisTitle;
	String yAxisTitle;
	AbstractSheetAdv sheet;
	
	NChartLegendPosition legendPosition;
	NChartGrouping grouping;
	NBarDirection direction;
	
	boolean threeD;
	
	public ChartImpl(AbstractSheetAdv sheet,String id,NChartType type,NViewAnchor anchor){
		this.sheet = sheet;
		this.id = id;
		this.type = type;
		this.anchor = anchor;
		this.data = createChartData(type);
		
		switch(type){//default initialization
		case BAR:
			direction = NBarDirection.HORIZONTAL;
			break;
		case COLUMN:
			direction = NBarDirection.VERTICAL;
			break;
		}
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
		switch(type){
		case AREA:
		case BAR:
		case COLUMN://same as bar
		case LINE:
		case DOUGHNUT://same as pie
		case PIE:
		case SCATTER://xy , reuse category
		case BUBBLE://xyz , reuse category
		case STOCK://stock, reuse category			
			return new GeneralChartDataImpl(this,id+"-data");
			
		//not supported	
		case OF_PIE:
		case RADAR:
		case SURFACE:
			return new UnsupportedChartDataImpl(this);
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
	@Override
	public void setLegendPosition(NChartLegendPosition pos) {
		this.legendPosition = pos;
	}
	@Override
	public NChartLegendPosition getLegendPosition() {
		return legendPosition;
	}
	@Override
	public void setGrouping(NChartGrouping grouping) {
		this.grouping = grouping;
	}
	@Override
	public NChartGrouping getGrouping() {
		return grouping;
	}
	@Override
	public NBarDirection getBarDirection() {
		return direction;
	}
	public void setBarDirection(NBarDirection direction){
		this.direction = direction;
	}
	@Override
	public boolean isThreeD() {
		return threeD;
	}
	@Override
	public void setThreeD(boolean threeD) {
		this.threeD = threeD;
	}
	
}
