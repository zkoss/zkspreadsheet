package org.zkoss.zss.ngmodel.impl.chart;

import org.zkoss.zss.ngmodel.NChart;

public class UnsupportedChartDataImpl extends ChartDataAdv {

	private static final long serialVersionUID = 1L;
	private NChart chart;
	public UnsupportedChartDataImpl(NChart chart){
		this.chart = chart;
	}

	@Override
	public NChart getChart() {
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
