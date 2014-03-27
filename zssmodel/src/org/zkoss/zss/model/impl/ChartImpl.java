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
package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.ViewAnchor;
import org.zkoss.zss.model.chart.SChartData;
import org.zkoss.zss.model.impl.chart.ChartDataAdv;
import org.zkoss.zss.model.impl.chart.GeneralChartDataImpl;
import org.zkoss.zss.model.impl.chart.UnsupportedChartDataImpl;
import org.zkoss.zss.model.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ChartImpl extends AbstractChartAdv {
	private static final long serialVersionUID = 1L;
	private String _id;
	private ChartType _type;
	private ViewAnchor _anchor;
	private ChartDataAdv _data;
	private String _title;
	private String _xAxisTitle;
	private String _yAxisTitle;
	private AbstractSheetAdv _sheet;
	
	private ChartLegendPosition _legendPosition;
	private ChartGrouping _grouping;
	private BarDirection _direction;
	
	private boolean _ThreeD;
	
	public ChartImpl(AbstractSheetAdv sheet,String id,ChartType type,ViewAnchor anchor){
		this._sheet = sheet;
		this._id = id;
		this._type = type;
		this._anchor = anchor;
		this._data = createChartData(type);
		
		Validations.argNotNull(anchor);
		
		switch(type){//default initialization
		case BAR:
			_direction = BarDirection.HORIZONTAL;
			break;
		case COLUMN:
			_direction = BarDirection.VERTICAL;
			break;
		}
	}
	@Override
	public SSheet getSheet(){
		checkOrphan();
		return _sheet;
	}
	@Override
	public String getId() {
		return _id;
	}
	@Override
	public ViewAnchor getAnchor() {
		return _anchor;
	}
	@Override
	public void setAnchor(ViewAnchor anchor) {
		Validations.argNotNull(anchor);
		this._anchor = anchor;		
	}
	@Override
	public ChartType getType(){
		return _type;
	}
	@Override
	public SChartData getData() {
		return _data;
	}
	@Override
	public String getTitle() {
		return _title;
	}
	@Override
	public void setTitle(String title) {
		this._title = title;
	}
	@Override
	public String getXAxisTitle() {
		return _xAxisTitle;
	}
	@Override
	public void setXAxisTitle(String xAxisTitle) {
		this._xAxisTitle = xAxisTitle;
	}
	@Override
	public String getYAxisTitle() {
		return _yAxisTitle;
	}
	@Override
	public void setYAxisTitle(String yAxisTitle) {
		this._yAxisTitle = yAxisTitle;
	}

	private ChartDataAdv createChartData(ChartType type){
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
			return new GeneralChartDataImpl(this,_id+"-data");
			
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
		
		_sheet = null;
	}
	@Override
	public void checkOrphan() {
		if (_sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	@Override
	public void setLegendPosition(ChartLegendPosition pos) {
		this._legendPosition = pos;
	}
	@Override
	public ChartLegendPosition getLegendPosition() {
		return _legendPosition;
	}
	@Override
	public void setGrouping(ChartGrouping grouping) {
		this._grouping = grouping;
	}
	@Override
	public ChartGrouping getGrouping() {
		return _grouping;
	}
	@Override
	public BarDirection getBarDirection() {
		return _direction;
	}
	public void setBarDirection(BarDirection direction){
		this._direction = direction;
	}
	@Override
	public boolean isThreeD() {
		return _ThreeD;
	}
	@Override
	public void setThreeD(boolean threeD) {
		this._ThreeD = threeD;
	}
	
}
