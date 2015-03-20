/* TableImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2014 6:56:44 PM, Created by henrichen

	Copyright (C) 2014 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SNamedStyle;
import org.zkoss.zss.model.STable;
import org.zkoss.zss.model.STableColumn;
import org.zkoss.zss.model.STableStyleInfo;

/**
 * @author henri
 * @since 3.8.0
 */
public class TableImpl implements STable {
	SAutoFilter _filter; // if _headerRowCount == 0; then _filter ==  null
	List<STableColumn> _columns;
	STableStyleInfo _tableStyleInfo;
	
//	SNamedStyle _totalsRowCellStyle;
//	SNamedStyle _dataCellStyle;
//	SNamedStyle _headerRowCellStyle;
//	
//	SCellStyle _totalsRowStyle; //Dxf
//	SCellStyle _dataStyle; //Dxf
//	SCellStyle _headerRowStyle; //Dxf
	
	int _totalsRowCount;
	int _headerRowCount;
	CellRegion _region;
	String _name;
	String _displayName;
	
	public TableImpl(String name, String displayName, 
			CellRegion region, 
			int headerRowCount, int totalsRowCount,
			STableStyleInfo info) {
//			SNamedStyle headerRowCellStyle, SCellStyle headerRowStyle,
//			SNamedStyle dataCellStyle, SCellStyle dataStyle,
//			SNamedStyle totalsRowCellStyle, SCellStyle totalsRowStyle) {
		_name = name;
		_displayName = displayName;
		_region = region;
		_headerRowCount = headerRowCount;
		_totalsRowCount = totalsRowCount;
		_tableStyleInfo = info;
		_columns  = new ArrayList<STableColumn>();
//		_headerRowCellStyle = headerRowCellStyle;
//		_totalsRowCellStyle = totalsRowCellStyle;
//		_dataCellStyle = dataCellStyle;
//		_headerRowStyle = headerRowStyle;
//		_totalsRowStyle = totalsRowStyle;
//		_dataStyle = dataStyle;
	}
	
	@Override
	public SAutoFilter getFilter() {
		// if no header row then there is no filter
		return _headerRowCount == 0 ? null : _filter;
	}
	public void enableFilter(boolean enable) {
		if ((_filter != null) == enable) return;
		
		if (enable) {
			if (getHeaderRowCount() == 0)
				setHeaderRowCount(1);
			else {
				final int l = _region.getColumn();
				final int t = _region.getRow();
				final int r = _region.getLastColumn();
				final int b = _region.getLastRow();
				final int tc = getTotalsRowCount();
				_filter = new AutoFilterImpl(new CellRegion(t, l, b - tc, r));
			}
		} else {
			_filter = null;
		}
	}

	@Override
	public List<STableColumn> getColumns() {
		return _columns;
	}
	@Override
	public void addColumn(STableColumn column) {
		_columns.add(column);
	}

	@Override
	public STableStyleInfo getTableStyleInfo() {
		return _tableStyleInfo;
	}
	public void setTableStyle(STableStyleInfo style) {
		_tableStyleInfo = style;
	}
//	@Override
//	public SNamedStyle getDataCellStyle() {
//		return _dataCellStyle;
//	}
//	@Override
//	public SNamedStyle getTotalsRowCellStyle() {
//		return _totalsRowCellStyle;
//	}
//	@Override
//	public SNamedStyle getHeaderRowCellStyle() {
//		return _headerRowCellStyle;
//	}
//	@Override
//	public SCellStyle getTotalsRowStyle() {
//		return _totalsRowStyle;
//	}
//	@Override
//	public SCellStyle getDataStyle() {
//		return _dataStyle;
//	}
//	@Override
//	public SCellStyle getHeaderRowStyle() {
//		return _headerRowStyle;
//	}
	@Override
	public int getTotalsRowCount() {
		return _totalsRowCount;
	}
	@Override
	public int getHeaderRowCount() {
		return _headerRowCount;
	}
	@Override
	public CellRegion getRegion() {
		return _region;
	}
	@Override
	public String getName() {
		return _name;
	}

	@Override
	public void setTotalsRowCount(int count) {
		if (_totalsRowCount == count) return;
		
		final int l = _region.getColumn();
		final int t = _region.getRow();
		final int r = _region.getLastColumn();
		final int b = _region.getLastRow();
		_region = new CellRegion(t, l, b + _totalsRowCount - count, r);
		_totalsRowCount = count;
	}

	@Override
	public void setHeaderRowCount(int count) {
		if (_headerRowCount == count) return;
		
		final int l = _region.getColumn();
		final int t = _region.getRow();
		final int r = _region.getLastColumn();
		final int b = _region.getLastRow();
		if (count == 0) {
			_filter = null;
		} else {
			final int tc = getTotalsRowCount();
			_filter = new AutoFilterImpl(new CellRegion(t + _headerRowCount - count, l, b - tc, r));
		}
		_region = new CellRegion(t + _headerRowCount - count, l, b, r);
		_headerRowCount = count;
	}

	@Override
	public void setName(String name) {
		_name = name;
	}

	@Override
	public String getDisplayName() {
		return _displayName;
	}

	@Override
	public void setDisplayName(String name) {
		_displayName = name;
	}
	
	@Override
	public CellRegion getNameRegion() {
		final int l = _region.getColumn();
		final int t = _region.getRow();
		final int r = _region.getLastColumn();
		final int b = _region.getLastRow();
		return new CellRegion(t + _headerRowCount, l, b - _totalsRowCount, r);
	}
	
	@Override
	public CellRegion getColumnRegion(String columnName) {
		int l = _region.getColumn();
		for (STableColumn tbCol : _columns) {
			if (columnName.equalsIgnoreCase(tbCol.getName())) {
				final int t = _region.getRow();
				final int r = _region.getLastColumn();
				final int b = _region.getLastRow();
				return new CellRegion(t + _headerRowCount, l, b - _totalsRowCount, l);
			}
			++l;
		}
		return null; 
	}
}
