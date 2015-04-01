/* STableColumn.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2014 2:25:58 PM, Created by henrichen

	Copyright (C) 2014 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

/**
 * Table column
 * @author henri
 * @since 3.8.0
 */
public interface STableColumn {
	String getName();
	void setName(String name);
	String getTotalsRowLabel();
	void setTotalsRowLabel(String label);
	STotalsRowFunction getTotalsRowFunction();
	void setTotalsRowFunction(STotalsRowFunction func);
	String getTotalsRowFormula();
	void setTotalsRowFormula(String formula);
	
	//the order is important which consist with Excel's model order
	public static enum STotalsRowFunction { 
		none,
		sum,
		min,
		max,
		average,
		count,
		countNums,
		stdDev,
		var,
		custom,
	}
}
