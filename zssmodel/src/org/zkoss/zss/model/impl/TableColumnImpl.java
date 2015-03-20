/* TableColumnImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2014 7:05:44 PM, Created by henrichen

	Copyright (C) 2014 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.STableColumn;

/**
 * Table column.
 * @author henri
 * @since 3.8.0
 */
public class TableColumnImpl implements STableColumn {
	String _name;
	SCellStyle _cellStyle;
	SCellStyle _dataStyle;
	SCellStyle _totalsRowStyle;
	String _totalsRowLabel;
	STotalsRowFunction _totalsRowFunction;
	
	public TableColumnImpl(String name) {
		_name = name;
	}
	@Override
	public String getName() {
		return _name;
	}

	@Override
	public SCellStyle getDataCellStyle() {
		return _cellStyle;
	}

	@Override
	public SCellStyle getDataStyle() {
		return _dataStyle;
	}

	@Override
	public SCellStyle getTotalsRowStyle() {
		return _totalsRowStyle;
	}

	@Override
	public String getTotalsRowLabel() {
		return _totalsRowLabel;
	}

	@Override
	public STotalsRowFunction getTotalsRowFunction() {
		return _totalsRowFunction;
	}

	@Override
	public void setName(String name) {
		_name = name;
	}

	@Override
	public void setDataCellStyle(SCellStyle style) {
		_cellStyle = style;
	}

	@Override
	public void setDataStyle(SCellStyle style) {
		_dataStyle = style;
	}

	@Override
	public void setTotalsRowStyle(SCellStyle style) {
		_totalsRowStyle = style;
	}

	@Override
	public void setTotalsRowLabel(String label) {
		_totalsRowLabel = label;
	}

	@Override
	public void setTotalsRowFunction(STotalsRowFunction func) {
		_totalsRowFunction = func;
	}
}
