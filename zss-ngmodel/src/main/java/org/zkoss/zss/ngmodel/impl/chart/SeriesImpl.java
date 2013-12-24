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
package org.zkoss.zss.ngmodel.impl.chart;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.AbstractChartAdv;
import org.zkoss.zss.ngmodel.impl.LinkedModelObject;
import org.zkoss.zss.ngmodel.impl.ObjectRefImpl;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class SeriesImpl implements NSeries,Serializable,LinkedModelObject{
	private static final long serialVersionUID = 1L;
	private FormulaExpression nameExpr;
	private FormulaExpression valueExpr;
	private FormulaExpression yValueExpr;
	private FormulaExpression zValueExpr;
	
	private AbstractChartAdv chart;
	private String id;
	
	private Object evalNameResult;
	private Object evalValuesResult;
	private Object evalYValuesResult;
	private Object evalZValuesResult;
	
	private boolean evaluated = false;
	
	/*package*/ void evalFormula(){
		if(!evaluated){
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			NSheet sheet = chart.getSheet();
			if(nameExpr!=null){
				EvaluationResult result = fe.evaluate(nameExpr,new FormulaEvaluationContext(sheet));

				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalNameResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalNameResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_VALUE);
				}
				
			}
			if(valueExpr!=null){
				EvaluationResult result = fe.evaluate(valueExpr,new FormulaEvaluationContext(sheet));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalValuesResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalValuesResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_VALUE);
				}
			}
			if(yValueExpr!=null){
				EvaluationResult result = fe.evaluate(yValueExpr,new FormulaEvaluationContext(sheet));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalYValuesResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalYValuesResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_VALUE);
				}
			}
			if(zValueExpr!=null){
				EvaluationResult result = fe.evaluate(zValueExpr,new FormulaEvaluationContext(sheet));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalZValuesResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalZValuesResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_VALUE);
				}
			}
			evaluated = true;
		}
	}
	
	public SeriesImpl(AbstractChartAdv chart,String id){
		this.chart = chart;
		this.id = id;
	}
	
	@Override
	public String getName() {
		evalFormula();
		return evalNameResult==null?null:(evalNameResult instanceof ErrorValue)?((ErrorValue)evalNameResult).getErrorString():evalNameResult.toString();
	}
	@Override
	public int getNumOfValue(){
		evalFormula();
		return ChartDataUtil.sizeOf(evalValuesResult);
	}
	@Override
	public Object getValue(int index) {
		evalFormula();
		if(index>=ChartDataUtil.sizeOf(evalValuesResult)){
			return null;
		}
		return ChartDataUtil.valueOf(evalValuesResult,index);
	}
	@Override
	public int getNumOfYValue(){
		evalFormula();
		return ChartDataUtil.sizeOf(evalYValuesResult);
	}
	@Override
	public Object getYValue(int index) {
		evalFormula();
		if(index>=ChartDataUtil.sizeOf(evalYValuesResult)){
			return null;
		}
		return ChartDataUtil.valueOf(evalYValuesResult,index);
	}
	public int getNumOfZValue(){
		evalFormula();
		return ChartDataUtil.sizeOf(evalZValuesResult);
	}
	@Override
	public Object getZValue(int index) {
		evalFormula();
		if(index>=ChartDataUtil.sizeOf(evalZValuesResult)){
			return null;
		}
		return ChartDataUtil.valueOf(evalZValuesResult,index);
	}
	
	@Override
	public void setFormula(String nameExpression,String valueExpression){
		setXYZFormula(nameExpression,valueExpression,null,null);
	}
	@Override
	public void setXYFormula(String nameExpression,String xValueExpression, String yValueExpression){
		setXYZFormula(nameExpression,xValueExpression,yValueExpression,null);
	}
	@Override
	public void setXYZFormula(String nameExpression,String xValueExpression, String yValueExpression,String zValueExpression){
		evaluated = false;
		clearFormulaDependency();
		
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		NSheet sheet = chart.getSheet();
		Ref ref = getRef();
		if(nameExpression!=null){
			nameExpr = fe.parse(nameExpression, new FormulaParseContext(sheet,ref));
		}else{
			nameExpr = null;
		}
		if(xValueExpression!=null){
			valueExpr = fe.parse(xValueExpression, new FormulaParseContext(sheet,ref));
		}else{
			valueExpr = null;
		}
		if(yValueExpression!=null){
			yValueExpr = fe.parse(yValueExpression, new FormulaParseContext(sheet,ref));
		}else{
			yValueExpr = null;
		}
		if(zValueExpression!=null){
			zValueExpr = fe.parse(zValueExpression, new FormulaParseContext(sheet,ref));
		}else{
			zValueExpr = null;
		}
	}
	
	@Override
	public boolean isFormulaParsingError() {
		boolean r = false;
		if(nameExpr!=null){
			r |= nameExpr.hasError();
		}
		if(!r && valueExpr!=null){
			r |= valueExpr.hasError();
		}
		if(!r && yValueExpr!=null){
			r |= yValueExpr.hasError();
		}
		if(!r && zValueExpr!=null){
			r |= zValueExpr.hasError();
		}
		
		return r;
	}

	@Override
	public String getNameFormula() {
		return nameExpr==null?null:nameExpr.getFormulaString();
	}

	@Override
	public String getValuesFormula() {
		return valueExpr==null?null:valueExpr.getFormulaString();
	}

	@Override
	public String getYValuesFormula() {
		return yValueExpr==null?null:yValueExpr.getFormulaString();
	}
	
	@Override
	public String getZValuesFormula() {
		return zValueExpr==null?null:zValueExpr.getFormulaString();
	}

	@Override
	public void clearFormulaResultCache() {
		evaluated = false;
		evalNameResult = evalValuesResult = evalYValuesResult = evalZValuesResult = null;		
	}
	
	private void clearFormulaDependency() {
		if(nameExpr!=null || valueExpr!=null || yValueExpr!=null || zValueExpr!=null){
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
		chart = null;
	}

	@Override
	public void checkOrphan() {
		if(chart==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public int getNumOfXValue() {
		return getNumOfValue();
	}

	@Override
	public Object getXValue(int index) {
		return getValue(index);
	}

	@Override
	public String getXValuesFormula() {
		return getValuesFormula();
	}

}
