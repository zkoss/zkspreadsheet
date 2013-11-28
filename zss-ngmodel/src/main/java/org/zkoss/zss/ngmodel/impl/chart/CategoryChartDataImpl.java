package org.zkoss.zss.ngmodel.impl.chart;

import java.io.Serializable;
import java.util.Collection;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.chart.NCategoryChartData;
import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.impl.AbstractChart;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;

public class CategoryChartDataImpl implements NCategoryChartData, Serializable{

	private static final long serialVersionUID = 1L;

	private FormulaExpression catFormula;
	
	private List<SeriesImpl> serieses = new LinkedList<SeriesImpl>();
	private AbstractChart chart;
	
	private Object evalResult;
	
	private boolean evaluated = false;
	
	
	public CategoryChartDataImpl(AbstractChart chart){
		this.chart = chart;
	}
	
	public NChart getChart(){
		return chart;
	}
	
	/*package*/ void evalFormula(){
		if(!evaluated){
			if(catFormula!=null){
				FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
				EvaluationResult result = fe.evaluate(catFormula,new FormulaEvaluationContext(chart.getSheet().getBook()));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_NAME);
				}
			}
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
		evalFormula();
		return ChartDataUtil.sizeOf(evalResult);
	}
	public Object getCategoryAt(int i) {
		evalFormula();
		if(i>=ChartDataUtil.sizeOf(evalResult)){
			return null;
		}
		return ChartDataUtil.valueOf(evalResult,i);
	}
	public NSeries addSeries() {
		SeriesImpl series = new SeriesImpl(chart);
		serieses.add(series);
		return series;
	}
	public void removeSeries(NSeries series) {
		serieses.remove(series);
	}
	public void setCategoriesFormula(String expr) {
		evaluated = false;
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		catFormula = fe.parse(expr, new FormulaParseContext(chart.getSheet().getBook()));
	}

	@Override
	public String getCategoriesFormula() {
		return catFormula==null?null:catFormula.getFormulaString();
	}
}
