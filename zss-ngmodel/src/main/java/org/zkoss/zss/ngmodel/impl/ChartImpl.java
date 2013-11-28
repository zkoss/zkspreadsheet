package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.chart.NChartData;
import org.zkoss.zss.ngmodel.impl.chart.CategoryChartDataImpl;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

public class ChartImpl extends AbstractChart {
	private static final long serialVersionUID = 1L;
	String id;
	NChartType type;
	NViewAnchor anchor;
	NChartData data;
	String title;
	String xAxisTitle;
	String yAxisTitle;
	AbstractSheet sheet;
	
	public ChartImpl(AbstractSheet sheet,String id,NChartType type,NViewAnchor anchor){
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

	private NChartData createChartData(NChartType type){
		//TODO for type;
		switch(type){
		case AREA:
		case BAR:
		case COLUMN:
		case LINE:
		case PIE:
			return new CategoryChartDataImpl(this);
		}
		throw new UnsupportedOperationException("unsupported chart type "+type);
	}

	@Override
	public void release() {
		checkOrphan();
		
		Ref ref = new RefImpl(this);
		((AbstractBookSeries)sheet.getBook().getBookSeries()).getDependencyTable().clearDependents(ref);
		
		sheet = null;
	}
	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
}
