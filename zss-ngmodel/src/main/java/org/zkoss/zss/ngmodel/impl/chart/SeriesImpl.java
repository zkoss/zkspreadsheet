package org.zkoss.zss.ngmodel.impl.chart;

import java.util.Collection;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.impl.AbstractBook;
import org.zkoss.zss.ngmodel.impl.AbstractChart;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;

public class SeriesImpl implements NSeries{
	
	private FormulaExpression nameExpr;
	private FormulaExpression valueExpr;
	private FormulaExpression yAxixExpr;
	
	private AbstractChart chart;
	
	private Object evalNameResult;
	private Object evalValueResult;
	private Object evalYValueResult;
	
	private boolean evaluated = false;
	
	/*package*/ void evalFormula(){
		if(!evaluated){
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			if(nameExpr!=null){
				EvaluationResult result = fe.evaluate(nameExpr,new FormulaEvaluationContext(chart.getSheet().getBook()));

				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalNameResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalNameResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_NAME);
				}
				
			}
			if(valueExpr!=null){
				EvaluationResult result = fe.evaluate(valueExpr,new FormulaEvaluationContext(chart.getSheet().getBook()));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalValueResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalValueResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_NAME);
				}
			}
			if(yAxixExpr!=null){
				EvaluationResult result = fe.evaluate(yAxixExpr,new FormulaEvaluationContext(chart.getSheet().getBook()));
				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalYValueResult = val;
				}else if(result.getType() == ResultType.ERROR){
					evalYValueResult = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_NAME);
				}
			}
			evaluated = true;
		}
	}
	
	public SeriesImpl(AbstractChart chart){
		this.chart = chart;
	}
	public String getName() {
		evalFormula();
		return evalNameResult==null?null:(evalNameResult instanceof ErrorValue)?((ErrorValue)evalNameResult).gettErrorString():evalNameResult.toString();
	}
	
	public int getNumOfValue(){
		evalFormula();
		return ChartDataUtil.sizeOf(evalValueResult);
	}
	
	public Object getValueAt(int index) {
		evalFormula();
		if(index>=ChartDataUtil.sizeOf(evalValueResult)){
			return null;
		}
		return ChartDataUtil.valueOf(evalValueResult,index);
	}
	
	public int getNumOfXValue(){
		return getNumOfValue();
	}
	public Object getXValueAt(int index) {
		return getValueAt(index);
	}
	
	public int getNumOfYValue(){
		evalFormula();
		return ChartDataUtil.sizeOf(evalYValueResult);
	}
	public Object getYValueAt(int index) {
		evalFormula();
		if(index>=ChartDataUtil.sizeOf(evalYValueResult)){
			return null;
		}
		return ChartDataUtil.valueOf(evalYValueResult,index);
	}
	public void setNameFormula(String expr) {
		evaluated = false;
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		nameExpr = fe.parse(expr, new FormulaParseContext(chart.getSheet().getBook()));
	}
	public void setValuesFormula(String expr) {
		evaluated = false;
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		valueExpr = fe.parse(expr, new FormulaParseContext(chart.getSheet().getBook()));
	}
	public void setXValuesFormula(String expr) {
		setValuesFormula(expr);
	}
	public void setYValuesFormula(String expr) {
		evaluated = false;
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		yAxixExpr = fe.parse(expr, new FormulaParseContext(chart.getSheet().getBook()));
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
	public String getXValuesFormula() {
		return getValuesFormula();
	}

	@Override
	public String getYValuesFormula() {
		return yAxixExpr==null?null:yAxixExpr.getFormulaString();
	}

}
