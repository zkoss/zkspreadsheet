/* CfValueObjectImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 12:50:53 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.sys.formula.FormulaExpression;

/**
 * @author henri
 * @since 3.8.2
 */
public class CFValueObjectImpl implements SCFValueObject, Serializable {
	private static final long serialVersionUID = 189125364763803062L;
	
	private CFValueObjectType type;
	private String value;
	private boolean gte = true; //default to true
	private FormulaExpression _formulaExpr;

	@Override
	public CFValueObjectType getType() {
		return type;
	}
	
	public void setType(CFValueObjectType type) {
		this.type = type;
	}

	@Override
	public String getValue() {
		return value;
	}
	
	public void setValue(String value) {
		this.value = value; 
	}

	@Override
	public boolean isGreaterOrEqual() {
		return gte;
	}
	
	public void setGreaterOrEqual(boolean b) {
		gte = b;
	}

	//ZSS-1142
	public CFValueObjectImpl cloneCFValueObject() {
		CFValueObjectImpl vo = new CFValueObjectImpl();
		vo.type = this.type;
		vo.value = this.value;
		vo.gte = this.gte;
		return vo;
	}
	
	//ZSS-1251
	public FormulaExpression getFormulaExpression() {
		return _formulaExpr;
	}
	
	//ZSS-1251
	public void setFormulaExpression(FormulaExpression expr) {
		_formulaExpr = expr;
		value = expr.getFormulaString();
	}
}
