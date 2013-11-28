package org.zkoss.zss.ngmodel.impl;

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
	
	public AbstractSheet getSheet(){
		checkOrphan();
		return sheet;
	}
	
	public String getId() {
		return id;
	}
	public NViewAnchor getAnchor() {
		return anchor;
	}
	public void setAnchor(NViewAnchor anchor) {
		this.anchor = anchor;
	}
	public NChartType getType(){
		return type;
	}
	public NChartData getData() {
		return data;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getXAxisTitle() {
		return xAxisTitle;
	}

	public void setXAxisTitle(String xAxisTitle) {
		this.xAxisTitle = xAxisTitle;
	}

	public String getYAxisTitle() {
		return yAxisTitle;
	}

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
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
}
