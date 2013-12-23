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
	/**
	 * Gets the num of x value, the result is same as {@link #getNumOfValue()}
	 * @return
	 */
	public int getNumOfXValue();
	/**
	 * Gets the x value, the result is same as {@link #getValue(int)}
	 * @param index
	 * @return
	 */
	public Object getXValue(int index);
	public int getNumOfYValue();
	public Object getYValue(int index);
	
	public int getNumOfZValue();
	public Object getZValue(int index);
	public void setFormula(String nameExpr,String valuesExpr);
	public void setXYFormula(String nameExpr,String xValuesExpr,String yValuesExpr);
	public void setXYZFormula(String nameExpr,String xValuesExpr,String yValuesExpr,String zValuesExpr);
	public String getNameFormula();
	public String getValuesFormula();
	/**
	 * Gets the x value formula, the result is same as {@link #getValuesFormula()}
	 * @return
	 */
	public String getXValuesFormula();
	public String getYValuesFormula();
	public String getZValuesFormula();
}
