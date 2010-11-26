/* FormulaMetaInfo.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 6:54:36 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.formula;


/**
 * @author Sam
 *
 */
public class FormulaMetaInfo {

	private String category;
	private String function;
	private String expression;
	private String description;
	private int requiredParameter;
	private boolean hasMultipleParameter;

	/**
	 * @param category
	 * @param function
	 * @param expression
	 * @param description
	 * @param requiredParameter
	 */
	public FormulaMetaInfo(String category, String function, String expression,
			String description, int requiredParameter, boolean hasMultipleParameter) {
		this.category = category;
		this.function = function;
		this.expression = expression;
		this.description = description;
		this.requiredParameter = requiredParameter;
		this.hasMultipleParameter = hasMultipleParameter;
	}
	public String getCategory() {
		return category;
	}
	public void setCategory(String category) {
		this.category = category;
	}
	public String getFunction() {
		return function;
	}
	public void setFunction(String function) {
		this.function = function;
	}
	public String getExpression() {
		return expression;
	}
	public void setExpression(String expression) {
		this.expression = expression;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public int getRequiredParameter() {
		return requiredParameter;
	}
	public void setRequiredParameter(int requiredParameter) {
		this.requiredParameter = requiredParameter;
	}
	
	
}
