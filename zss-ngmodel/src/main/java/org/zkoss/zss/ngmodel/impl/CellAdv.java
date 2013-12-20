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
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Date;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NRichText;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class CellAdv implements NCell,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;
	
	protected void checkType(CellType... types){
		Set<CellType> set = new LinkedHashSet<CellType>();
		for(CellType t:types){
			set.add(t);
		}
		
		if(!set.contains(getType())){
			throw new IllegalStateException("is "+getType()+", not the one of "+types);
		}
	}
	protected void checkFormulaResultType(CellType... types){
		if(!getType().equals(CellType.FORMULA)){
			throw new IllegalStateException("is "+getType()+", not the one of "+types);
		}
		
		Set<CellType> set = new LinkedHashSet<CellType>();
		for(CellType t:types){
			set.add(t);
		}
		if(!set.contains(getFormulaResultType())){
			throw new IllegalStateException("is "+getFormulaResultType()+", not the one of "+types);
		}
	}
	
	/*package*/ abstract void evalFormula();
	/*package*/ abstract Object getValue(boolean evaluatedVal);
	/*package*/ abstract NCellStyle getCellStyle(boolean local);
	
	@Override
	public Object getValue(){
		return getValue(true);
	}
	
	@Override
	public void setStringValue(String value) {
		setValue(value);
	}

	@Override
	public String getStringValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.STRING,CellType.BLANK);
		}else{
			checkType(CellType.STRING,CellType.BLANK);
		}
		Object val = getValue();
		return val==null?"":(String)val;
	}

	@Override
	public void setNumberValue(Double number) {
		setValue(number);
	}

	@Override
	public Double getNumberValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.NUMBER,CellType.BLANK);
		}else{
			checkType(CellType.NUMBER,CellType.BLANK);
		}
		Object val = getValue();
		return val==null?0.0:(Double)getValue();
	}

	@Override
	public void setDateValue(Date date) {
		double num = EngineFactory.getInstance().getCalendarUtil().dateToDoubleValue(date, false);
		setNumberValue(num);
	}

	@Override
	public Date getDateValue() {
		Number num = getNumberValue();
		return EngineFactory.getInstance().getCalendarUtil().doubleValueToDate(num.doubleValue(), false);
	}
	
	@Override
	public void setBooleanValue(Boolean date) {
		setValue(date);
	}

	@Override
	public Boolean getBooleanValue() {
		CellType type = getType();
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.BOOLEAN,CellType.BLANK);
		}else{
			checkType(CellType.BOOLEAN,CellType.BLANK);
		}
		
		return Boolean.TRUE.equals(getValue());
	}

	@Override
	public ErrorValue getErrorValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.ERROR);
		}else{
			checkType(CellType.ERROR);
		}
		return (ErrorValue)getValue();
	}

	@Override
	public void setErrorValue(ErrorValue errorValue) {
		setValue(errorValue);
	}
	
	@Override
	public void setFormulaValue(String formula) {
		checkOrphan();
		Validations.argNotNull(formula);
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		FormulaExpression expr = fe.parse(formula, new FormulaParseContext(this ,new RefImpl(this)));
		setValue(expr);
	}
	

	@Override
	public String getFormulaValue() {
		checkType(CellType.FORMULA);
		FormulaExpression expr = (FormulaExpression)getValue(false);
		return expr.getFormulaString();
	}
	

	@Override
	public NRichText setupRichText() {
		NRichText text = new RichTextImpl();
		setRichText(text);
		return text;
	}
	@Override
	public NHyperlink setupHyperlink(){
		NHyperlink hyperlink = new HyperlinkImpl();
		setHyperlink(hyperlink);
		return hyperlink;
	}
	@Override
	public NComment setupComment(){
		NComment comment = new CommentImpl();
		setComment(comment);
		return comment;
	}
}
