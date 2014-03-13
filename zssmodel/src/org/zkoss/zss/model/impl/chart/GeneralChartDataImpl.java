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
import org.zkoss.zss.model.InvalidModelOpException;
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

	private FormulaExpression _catFormula;
	
	final private List<SeriesImpl> _serieses = new LinkedList<SeriesImpl>();
	private AbstractChartAdv _chart;
	final private String _id;
	
	private Object _evalResult;
	
	private boolean _evaluated = false;
	
	private int _seriesCount = 0;
	
	public GeneralChartDataImpl(AbstractChartAdv chart,String id){
		this._chart = chart;
		this._id = id;
	}
	
	public SChart getChart(){
		return _chart;
	}
	
	/*package*/ void evalFormula(){
		if(!_evaluated){
			if(_catFormula!=null){
				FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
				EvaluationResult result = fe.evaluate(_catFormula,new FormulaEvaluationContext(_chart.getSheet()));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					_evalResult = val;
				}else if(result.getType() == ResultType.ERROR){
					_evalResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_NAME);
				}
			}
			//TODO
			_evaluated = true;
		}
	}
	
	public int getNumOfSeries() {
		return _serieses.size();
	}
	public SSeries getSeries(int i) {
		return _serieses.get(i);
	}
	public int getNumOfCategory() {
		evalFormula();
		return EvaluationUtil.sizeOf(_evalResult);
	}
	public Object getCategory(int i) {
		evalFormula();
		if(i>=EvaluationUtil.sizeOf(_evalResult)){
			return null;
		}
		return EvaluationUtil.valueOf(_evalResult,i);
	}
	public SSeries addSeries() {
		SeriesImpl series = new SeriesImpl(_chart,_chart.getId() + "-series" + (_seriesCount++));
		_serieses.add(series);
		return series;
	}
	
	protected void checkOwnership(SSeries series){
		if(!_serieses.contains(series)){
			throw new IllegalStateException("doesn't has ownership "+ series);
		}
	}
	
	public void removeSeries(SSeries series) {
		checkOwnership(series);
		((SeriesImpl)series).destroy();
		_serieses.remove(series);
	}
	public void setCategoriesFormula(String expr) {
		checkOrphan();
		_evaluated = false;
		
		clearFormulaDependency();
		
		if(expr!=null){
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			_catFormula = fe.parse(expr, new FormulaParseContext(_chart.getSheet(),getRef()));
		}else{
			_catFormula = null;
		}
	}

	@Override
	public String getCategoriesFormula() {
		return _catFormula==null?null:_catFormula.getFormulaString();
	}

	@Override
	public void clearFormulaResultCache() {
		_evalResult = null;
		_evaluated = false;
		for(SeriesImpl series:_serieses){
			series.clearFormulaResultCache();
		}
	}
	
	@Override
	public boolean isFormulaParsingError() {
		return _catFormula==null?false:_catFormula.hasError();
	}
	
	private void clearFormulaDependency(){
		if(_catFormula!=null){
			((AbstractBookSeriesAdv) _chart.getSheet().getBook().getBookSeries())
					.getDependencyTable().clearDependents(getRef());
		}
	}
	
	private Ref getRef(){
		return new ObjectRefImpl(_chart,new String[]{_chart.getId(),_id});
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency();
		clearFormulaResultCache();
		for(SeriesImpl series:_serieses){
			series.destroy();
		}
		_chart = null;
	}

	@Override
	public void checkOrphan() {
		if(_chart==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

}
