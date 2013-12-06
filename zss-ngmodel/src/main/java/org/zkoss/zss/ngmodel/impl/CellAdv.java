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

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NRichText;
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
	
	abstract protected void evalFormula();
	abstract protected Object getValue(boolean valueOfFormula);
	
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
			checkFormulaResultType(CellType.STRING);
		}else{
			checkType(CellType.STRING);
		}
		return (String)getValue();
	}
	
	@Override
	public void setRichTextValue(NRichText value) {
		setValue(value);
	}

	@Override
	public NRichText getRichTextValue() {
		checkType(CellType.RICHTEXT);
		return (NRichText)getValue();
	}	

	@Override
	public void setNumberValue(Number number) {
		setValue(number);
	}

	@Override
	public Number getNumberValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.NUMBER);
		}else{
			checkType(CellType.NUMBER);
		}
		return (Number)getValue();
	}

	@Override
	public void setDateValue(Date date) {
		setValue(date);
	}

	@Override
	public Date getDateValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.DATE);
		}else{
			checkType(CellType.DATE);
		}
		return (Date)getValue();
	}
	
	@Override
	public void setBooleanValue(Boolean date) {
		setValue(date);
	}

	@Override
	public Boolean getBooleanValue() {
		if(getType() == CellType.FORMULA){
			evalFormula();
			checkFormulaResultType(CellType.BOOLEAN);
		}else{
			checkType(CellType.BOOLEAN);
		}
		return (Boolean)getValue();
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
		if(formula.startsWith("=")){
			formula = formula.substring(1);
		}
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		FormulaExpression expr = fe.parse(formula, new FormulaParseContext(getSheet().getBook()));
		setValue(expr);
	}

	@Override
	public String getFormulaValue() {
		checkType(CellType.FORMULA);
		FormulaExpression expr = (FormulaExpression)getValue(false);
		return expr.getFormulaString();
	}
	

	@Override
	public NRichText setRichTextValue() {
		NRichText text = new RichTextImpl();
		setRichTextValue(text);
		return text;
	}
	@Override
	public NHyperlink setHyperlink(){
		NHyperlink hyperlink = new HyperlinkImpl();
		setHyperlink(hyperlink);
		return hyperlink;
	}
	@Override
	public NComment setComment(){
		NComment comment = new CommentImpl();
		setComment(comment);
		return comment;
	}
}
