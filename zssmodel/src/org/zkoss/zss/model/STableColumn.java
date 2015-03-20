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
	SCellStyle getDataCellStyle(); // dataCellStyle - by name
	void setDataCellStyle(SCellStyle style);
	SCellStyle getDataStyle(); // dataDxfId
	void setDataStyle(SCellStyle style);
	SCellStyle getTotalsRowStyle(); // totalsRowDxfId
	void setTotalsRowStyle(SCellStyle style);
	String getTotalsRowLabel();
	void setTotalsRowLabel(String label);
	STotalsRowFunction getTotalsRowFunction();
	void setTotalsRowFunction(STotalsRowFunction func);
	
	public static enum STotalsRowFunction {
		AVERAGE,
		COUNT,
		COUNT_NUMS,
		CUSTOM,
		MAX,
		MIN,
		NONE,
		STD_DEV,
		SUM,
		VAR;
	}
}
