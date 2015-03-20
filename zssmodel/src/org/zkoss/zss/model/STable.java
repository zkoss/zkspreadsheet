/* STable.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2014 2:24:34 PM, Created by henrichen

	Copyright (C) 2014 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;

/**
 * Table
 * @author henri
 * @since 3.8.0
 */
public interface STable {
	SAutoFilter getFilter(); // turn off header Row; then this is null; ref and region is the same; then no totals row
	void enableFilter(boolean enable);
	
	void addColumn(STableColumn column);
	List<STableColumn> getColumns();
	STableStyleInfo getTableStyleInfo();
	
//	SNamedStyle getDataCellStyle(); // dataCellStyle
//	SNamedStyle getTotalsRowCellStyle(); // totalsRowCellStyle
//	SNamedStyle getHeaderRowCellStyle(); // headerRowCellStyle
//	SCellStyle getTotalsRowStyle(); //totalsRowDxfId
//	SCellStyle getDataStyle(); //dataDxfId
//	SCellStyle getHeaderRowStyle(); //headerRowDxfId
	int getTotalsRowCount(); //totalsRowCount; 0 if not show; default is 0
	void setTotalsRowCount(int count);
	int getHeaderRowCount(); //headerRowCount; 0 if not show; default is 1
	void setHeaderRowCount(int count);
	CellRegion getRegion(); // ref
	String getName(); //name (Name used in formula)
	void setName(String name);
	String getDisplayName();
	void setDisplayName(String name);
	
	CellRegion getNameRegion(); // associated Name region
	CellRegion getColumnRegion(String columnName);
}
