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
package org.zkoss.zss.ngmodel.chart;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NGeneralChartData extends NChartData{

	/**
	 * Return formula parsing state.
	 * @return true if has error, false if no error or no formula
	 */
	public boolean isFormulaParsingError();
	
	public int getNumOfSeries();
	public NSeries getSeries(int i);
	
	public int getNumOfCategory();
	public Object getCategory(int i);
	
	
	
	public NSeries addSeries();
	public void removeSeries(NSeries series);
	
	public String getCategoriesFormula();
	public void setCategoriesFormula(String expr);
	
}
