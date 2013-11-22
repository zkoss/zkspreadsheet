package org.zkoss.zss.ngmodel.impl;

import java.util.Date;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.Validations;

public abstract class AbstractCell implements NCell{

	void release() {
	}
	protected void checkType(CellType type){
		if(!getType().equals(type)){
			throw new IllegalStateException("is "+getType()+", not the "+type);
		}
	}
	protected void checkFormulaResultType(CellType type){
		checkType(CellType.FORMULA);
		if(!getFormulaResultType().equals(type)){
			throw new IllegalStateException("formula result is "+getFormulaResultType()+", not the "+type);
		}
	}
	abstract void evalFormula();
	abstract public void setValue(Object value);
	abstract protected Object getValue(boolean eval);
	
	public Object getValue(){
		return getValue(true);
	}
	
	public void setStringValue(String value) {
		setValue(value);
	}

	public String getStringValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.STRING);
		}else{
			checkType(CellType.STRING);
		}
		return (String)getValue();
	}

	public void setNumberValue(Number number) {
		setValue(number);
	}

	public Number getNumberValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.NUMBER);
		}else{
			checkType(CellType.NUMBER);
		}
		return (Number)getValue();
	}

	public void setDateValue(Date date) {
		setValue(date);
	}

	public Date getDateValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.DATE);
		}else{
			checkType(CellType.DATE);
		}
		return (Date)getValue();
	}

	public ErrorValue getErrorValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.ERROR);
		}else{
			checkType(CellType.ERROR);
		}
		return (ErrorValue)getValue();
	}

	public void setErrorValue(ErrorValue errorValue) {
		setValue(errorValue);
	}
	
	
	public void setFormulaValue(String formula) {
		Validations.argNotNull(formula);
		if(formula.startsWith("=")){
			formula = formula.substring(1);
		}
		FormulaEngine fe = EngineFactory.getInstance().getFormulaEngine();
		FormulaExpression expr = fe.parse(formula, new FormulaParseContext());
		setValue(expr);
	}

	public String getFormulaValue() {
		checkType(CellType.FORMULA);
		FormulaExpression expr = (FormulaExpression)getValue(false);
		return expr.getFormulaString();
	}
}
