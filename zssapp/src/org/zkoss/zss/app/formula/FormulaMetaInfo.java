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

import java.util.HashMap;

/**
 * @author Sam
 *
 */
public class FormulaMetaInfo {

	private String category;
	private String function;
	private String display;
	private String description;
	private int requiredParameter;

	/**
	 * @param category
	 * @param function
	 * @param display
	 * @param description
	 * @param requiredParameter
	 */
	public FormulaMetaInfo(String category, String function, String display,
			String description, int requiredParameter) {
		this.category = category;
		this.function = function;
		this.display = display;
		this.description = description;
		this.requiredParameter = requiredParameter;
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
	public String getDisplay() {
		return display;
	}
	public void setDisplay(String display) {
		this.display = display;
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
