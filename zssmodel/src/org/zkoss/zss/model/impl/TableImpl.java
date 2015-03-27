/* TableImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 9, 2014 6:56:44 PM, Created by henrichen

	Copyright (C) 2014 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.ptg.TablePtg;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SNamedStyle;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.STable;
import org.zkoss.zss.model.STableColumn;
import org.zkoss.zss.model.STableStyleInfo;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;

/**
 * @author henri
 * @since 3.8.0
 */
public class TableImpl implements STable,LinkedModelObject,Serializable{
	AbstractBookAdv _book;
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
	SheetRegion _region;
	String _name;
	String _displayName;
	
	public TableImpl(AbstractBookAdv book, String name, String displayName, 
			SheetRegion region, 
			int headerRowCount, int totalsRowCount,
			STableStyleInfo info) {
//			SNamedStyle headerRowCellStyle, SCellStyle headerRowStyle,
//			SNamedStyle dataCellStyle, SCellStyle dataStyle,
//			SNamedStyle totalsRowCellStyle, SCellStyle totalsRowStyle) {
		_book = book;
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

	//ZSS-967
	@Override
	public SBook getBook() {
		return _book;
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
	public SheetRegion getAllRegion() {
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
		_region = new SheetRegion(_region.getSheet(), new CellRegion(t, l, b + _totalsRowCount - count, r));
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
		_region = new SheetRegion(_region.getSheet(), new CellRegion(t + _headerRowCount - count, l, b, r));
		_headerRowCount = count;
	}

	@Override
	public void setName(String newname) {
		checkOrphan();
		_name = newname;
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
	public SheetRegion getDataRegion() {
		final int l = _region.getColumn();
		final int t = _region.getRow();
		final int r = _region.getLastColumn();
		final int b = _region.getLastRow();
		return new SheetRegion(_region.getSheet(), new CellRegion(t + _headerRowCount, l, b - _totalsRowCount, r));
	}
	
	@Override
	public SheetRegion getColumnsRegion(String columnName1, String columnName2) {
		if (columnName1 == null) {
			if (columnName2 == null)
				return null;
			else {
				columnName1 = columnName2;
				columnName2 = null;
			}
		}
		int c1 = -1;
		int c2 = -1;
		int l = _region.getColumn();
		SSheet sheet  = _region.getSheet();
		for (STableColumn tbCol : _columns) {
			if (columnName1.equalsIgnoreCase(tbCol.getName())) {
				c1 = l;
				if (columnName2 == null || c2 >= 0) break;
			}
			if (tbCol.getName().equalsIgnoreCase(columnName2)) {
				c2 = l;
				if (c1 >= 0) break;
			}
			++l;
		}
		if (c1 < 0 || (columnName2 != null && c2 < 0))
			return null;
		
		if (c2 < 0) c2 = c1;
		final int t = _region.getRow();
		final int b = _region.getLastRow();
		if (c2 < c1) {
			final int tmp = c1;
			c1 = c2;
			c2 = tmp;
		}
		return new SheetRegion(sheet, new CellRegion(t + _headerRowCount, c1, b - _totalsRowCount, c2));
	}

	@Override
	public SheetRegion getHeadersRegion() {
		if (_headerRowCount == 0) return null; // no headers row at all
		final int l = _region.getColumn();
		final int t = _region.getRow();
		final int r = _region.getLastColumn();
		return new SheetRegion(_region.getSheet(), new CellRegion(t, l, t, r));
	}

	@Override
	public SheetRegion getTotalsRegion() {
		if (_totalsRowCount == 0) return null; //no totals row at all
		final int l = _region.getColumn();
		final int r = _region.getLastColumn();
		final int b = _region.getLastRow();
		return new SheetRegion(_region.getSheet(), new CellRegion(b, l, b, r));
	}

	@Override
	public SheetRegion getThisRowRegion(int rowIdx) {
		final int t = _region.getRow() + _headerRowCount;
		final int b = _region.getLastRow() - _totalsRowCount;
		if (t > rowIdx || rowIdx > b) {
			throw new IndexOutOfBoundsException("expect rowIdx(" + rowIdx + ") is between "+ t + " and " + b);
		}
		final int l = _region.getColumn();
		final int r = _region.getLastColumn();
		return new SheetRegion(_region.getSheet(), new CellRegion(rowIdx, l, rowIdx, r));
	}

	@Override
	public SheetRegion getItemRegion(TablePtg.Item item, int rowIdx) {
		if (item == null)
			return null;
		
		switch(item) {
		case ALL:
			return getAllRegion();
		case DATA:
			return getDataRegion();
		case HEADERS:
			return getHeadersRegion();
		case TOTALS:
			return getTotalsRegion();
		case THIS_ROW:
			return getThisRowRegion(rowIdx);
		}
		return null;
	}
	
	@Override
	public void destroy() {
		checkOrphan();
		_book = null;
	}

	@Override
	public void checkOrphan() {
		if(_book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	//ZSS-967
	@Override
	public STableColumn getColumnAt(int colIdx) {
		CellRegion rgn = _region.getRegion();
		final int idx = colIdx - rgn.getColumn();
		return idx < 0 || idx >= _columns.size() ? null : _columns.get(idx);
	}
}