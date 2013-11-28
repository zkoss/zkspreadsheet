package org.zkoss.zss.ngmodel.impl.chart;

import java.io.Serializable;
import java.util.Collection;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.chart.NCategoryChartData;
import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.impl.AbstractChart;

public class CategoryChartDataImpl implements NCategoryChartData, Serializable{

	private String categoriesExpr;
	
	private List<SeriesImpl> serieses = new LinkedList<SeriesImpl>();
	private AbstractChart chart;
	
	private Object evaluatedResult;
	
	private boolean evaluated = false;
	
	
	public CategoryChartDataImpl(AbstractChart chart){
		this.chart = chart;
	}
	
	public NChart getChart(){
		return chart;
	}
	
	/*package*/ void eval(){
		if(!evaluated){
			//TODO
			evaluated = true;
		}
	}
	
	public int getNumOfSeries() {
		return serieses.size();
	}
	public NSeries getSeriesAt(int i) {
		return serieses.get(i);
	}
	public int getNumOfCategory() {
		eval();
		return sizeOf(evaluatedResult);
	}
	public Object getCategoryAt(int i) {
		eval();
		return valueOf(evaluatedResult,i);
	}
	public NSeries addSeries() {
		SeriesImpl series = new SeriesImpl(chart);
		serieses.add(series);
		return series;
	}
	public void removeSeries(NSeries series) {
		serieses.remove(series);
	}
	public void setCategoriesExpression(String expr) {
		categoriesExpr = expr;
		evaluated = false;
	}
	
	static final int sizeOf(Object obj){
		if(obj==null){
			return 0;
		}else if(obj instanceof Collection){
			return ((Collection)obj).size();
		}else if(obj.getClass().isArray()){
			return ((Object[])obj).length;
		}else{
			return 1;
		}
	}
	
	static final Object valueOf(Object obj,int index){
		if(obj instanceof List){
			return ((List)obj).get(index);
		}else if(obj instanceof Collection){
			return ((Collection)obj).toArray()[index];
		}else if(obj.getClass().isArray()){
			return ((Object[])obj)[index];
		}else{
			return obj;
		}
	}


}
