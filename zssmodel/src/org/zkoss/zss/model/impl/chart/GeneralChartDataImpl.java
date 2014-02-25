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
package org.zkoss.zss.model.impl.chart;

import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.InvalidateModelOpException;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.chart.SGeneralChartData;
import org.zkoss.zss.model.chart.SSeries;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.impl.AbstractChartAdv;
import org.zkoss.zss.model.impl.EvaluationUtil;
import org.zkoss.zss.model.impl.ObjectRefImpl;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class GeneralChartDataImpl extends ChartDataAdv implements SGeneralChartData{

	private static final long serialVersionUID = 1L;

	private FormulaExpression catFormula;
	
	final private List<SeriesImpl> serieses = new LinkedList<SeriesImpl>();
	private AbstractChartAdv chart;
	final private String id;
	
	private Object evalResult;
	
	private boolean evaluated = false;
	
	private int seriesCount = 0;
	
	public GeneralChartDataImpl(AbstractChartAdv chart,String id){
		this.chart = chart;
		this.id = id;
	}
	
	public SChart getChart(){
		return chart;
	}
	
	/*package*/ void evalFormula(){
		if(!evaluated){
			if(catFormula!=null){
				FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
				EvaluationResult result = fe.evaluate(catFormula,new FormulaEvaluationContext(chart.getSheet()));
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
	public SSeries getSeries(int i) {
		return serieses.get(i);
	}
	public int getNumOfCategory() {
		evalFormula();
		return EvaluationUtil.sizeOf(evalResult);
	}
	public Object getCategory(int i) {
		evalFormula();
		if(i>=EvaluationUtil.sizeOf(evalResult)){
			return null;
		}
		return EvaluationUtil.valueOf(evalResult,i);
	}
	public SSeries addSeries() {
		SeriesImpl series = new SeriesImpl(chart,chart.getId() + "-series" + (seriesCount++));
		serieses.add(series);
		return series;
	}
	
	protected void checkOwnership(SSeries series){
		if(!serieses.contains(series)){
			throw new IllegalStateException("doesn't has ownership "+ series);
		}
	}
	
	public void removeSeries(SSeries series) {
		checkOwnership(series);
		((SeriesImpl)series).destroy();
		serieses.remove(series);
	}
	public void setCategoriesFormula(String expr) {
		checkOrphan();
		evaluated = false;
		
		clearFormulaDependency();
		
		//TODO dependency tracking on chart
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		catFormula = fe.parse(expr, new FormulaParseContext(chart.getSheet(),getRef()));
	}

	@Override
	public String getCategoriesFormula() {
		return catFormula==null?null:catFormula.getFormulaString();
	}

	@Override
	public void clearFormulaResultCache() {
		evalResult = null;
		evaluated = false;
		for(SeriesImpl series:serieses){
			series.clearFormulaResultCache();
		}
	}
	
	@Override
	public boolean isFormulaParsingError() {
		return catFormula==null?false:catFormula.hasError();
	}
	
	private void clearFormulaDependency(){
		if(catFormula!=null){
			((AbstractBookSeriesAdv) chart.getSheet().getBook().getBookSeries())
					.getDependencyTable().clearDependents(getRef());
		}
	}
	
	private Ref getRef(){
		return new ObjectRefImpl(chart,new String[]{chart.getId(),id});
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency();
		clearFormulaResultCache();
		for(SeriesImpl series:serieses){
			series.destroy();
		}
		chart = null;
	}

	@Override
	public void checkOrphan() {
		if(chart==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

}
