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

import org.zkoss.zss.ngmodel.FormulaContent;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NSeries extends FormulaContent{

	public String getName();
	
	public int getNumOfValue();
	public Object getValue(int index);
	
	//for Scatter, xy chart
	public int getNumOfYValue();
	public Object getYValue(int index);
	public void setFormula(String nameExpr,String valuesExpr);
	public void setFormula(String nameExpr,String xValuesExpr,String yValuesExpr);
	public String getNameFormula();
	public String getValuesFormula();
	public String getYValuesFormula();
}
