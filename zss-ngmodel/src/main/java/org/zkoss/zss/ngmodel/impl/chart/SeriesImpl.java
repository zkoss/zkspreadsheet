package org.zkoss.zss.ngmodel.impl.chart;

import java.util.Collection;

import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.impl.AbstractBook;
import org.zkoss.zss.ngmodel.impl.AbstractChart;

public class SeriesImpl implements NSeries{
	
	String nameExpr;
	String valueExpr;
	String yAxixExpr;
	
	private AbstractChart chart;
	
	private Object evaluatedNameResult;
	private Object evaluatedValueResult;
	private Object evaluatedYValueResult;
	
	private boolean evaluated = false;
	
	/*package*/ void eval(){
		if(!evaluated){
		
			evaluated = true;
		}
	}
	
	public SeriesImpl(AbstractChart chart){
		this.chart = chart;
	}
	public String getName() {
		eval();
		return evaluatedNameResult==null?null:evaluatedNameResult.toString();
	}
	
	public int getNumOfValue(){
		eval();
		return CategoryChartDataImpl.sizeOf(evaluatedNameResult);
	}
	
	public Object getValueAt(int index) {
		// TODO Auto-generated method stub
		return null;
	}
	public Object getXValueAt(int index) {
		return getValueAt(index);
	}
	public Object getYValueAt(int index) {
		// TODO Auto-generated method stub
		return null;
	}
	public void setNameExpression(String expr) {
		// TODO Auto-generated method stub
		
	}
	public void setValuesExpresion(String expr) {
		// TODO Auto-generated method stub
		
	}
	public void setXValuesExpression(String expr) {
		// TODO Auto-generated method stub
		
	}
	public void setYValuesExpression(String expr) {
		// TODO Auto-generated method stub
		
	}

}
