package org.zkoss.zss.model.impl.chart;

import org.zkoss.zss.model.SChart;

public class UnsupportedChartDataImpl extends ChartDataAdv {

	private static final long serialVersionUID = 1L;
	private SChart chart;
	public UnsupportedChartDataImpl(SChart chart){
		this.chart = chart;
	}

	@Override
	public SChart getChart() {
		return chart;
	}

	@Override
	public void clearFormulaResultCache() {
	}

	@Override
	public boolean isFormulaParsingError() {
		return false;
	}

	@Override
	public void destroy() {
		chart = null;
	}

	@Override
	public void checkOrphan() {
		if(chart==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

}
