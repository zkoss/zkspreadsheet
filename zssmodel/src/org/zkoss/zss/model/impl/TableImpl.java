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

import org.zkoss.poi.ss.formula.ptg.TablePtg;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBorder;
import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.SBorderLine;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.STableColumn;
import org.zkoss.zss.model.STableStyle;
import org.zkoss.zss.model.STableStyleElem;
import org.zkoss.zss.model.STableStyleInfo;
import org.zkoss.zss.model.SheetRegion;

/**
 * @author henri
 * @since 3.8.0
 */
public class TableImpl extends AbstractTableAdv implements LinkedModelObject {
	private static final long serialVersionUID = 1L;
	
	AbstractBookAdv _book;
	SAutoFilter _filter; // if _headerRowCount == 0; then _filter ==  null
	List<STableColumn> _columns;
	STableStyleInfo _tableStyleInfo;
	
	int _totalsRowCount;
	int _headerRowCount;
	SheetRegion _region;
	String _name;
	String _displayName;
	
	public TableImpl(AbstractBookAdv book, String name, String displayName, 
			SheetRegion region, 
			int headerRowCount, int totalsRowCount,
			STableStyleInfo info) {
		_book = book;
		_name = name;
		_displayName = displayName;
		_region = region;
		_headerRowCount = headerRowCount;
		_totalsRowCount = totalsRowCount;
		_tableStyleInfo = info;
		_columns  = new ArrayList<STableColumn>();
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

	//ZSS-977
	private static class CellStylePicker {
		private SFill fill = null;
		private SFont font = null;
		private SBorderLine left = null;
		private SBorderLine top = null;
		private SBorderLine right = null;
		private SBorderLine bottom = null;
	
		//for lastCol/firstCol/colStripe1/colStripe2/rowStripe1/rowStripe2
		private boolean pickDataStyle(SCellStyle style, SCellStyle totalsRowStyle) {
			if (style != null) {
				final SBorder nextBorder0 = 
						totalsRowStyle != null ? totalsRowStyle.getBorder() : null;
				final SBorderLine bottom02 = 
						nextBorder0 != null ? nextBorder0.getTopLine() : null; 
				final SFont font0 = style.getFont();
				final SFill fill0 = style.getFill();
				final SBorder border0 = style.getBorder();
				if (border0 != null) {
					final SBorderLine left0 = border0.getLeftLine();
					final SBorderLine top0 = border0.getTopLine();
					final SBorderLine right0 = border0.getRightLine();
					final SBorderLine bottom0 = border0.getBottomLine();
					return pickStyle(font0, fill0, left0, top0, right0, bottom02 != null ? bottom02 : bottom0);
				} else {
					return pickStyle(font0, fill0, null, null, null, bottom02);
				}
			}
			return false;
		}

		//for wholeTable
		private boolean pickWholeTableStyle(SCellStyle style, SCellStyle totalsRowStyle, 
				boolean firstRow, boolean firstCol, boolean lastRow, boolean lastCol) {
			if (style != null) {
				final SBorder nextBorder0 = 
						totalsRowStyle != null ? totalsRowStyle.getBorder() : null;
				final SBorderLine bottom02 = 
						nextBorder0 != null ? nextBorder0.getTopLine() : null; 
				final SFont font0 = style.getFont();
				final SFill fill0 = style.getFill();
				final SBorder border0 = style.getBorder();
				if (border0 != null) {
					final SBorderLine left0 = firstCol ? border0.getLeftLine() : border0.getVerticalLine();
					final SBorderLine right0 = lastCol ? border0.getRightLine() : border0.getVerticalLine();
					final SBorderLine top0 = firstRow ? border0.getTopLine() : border0.getHorizontalLine();
					final SBorderLine bottom0 = lastRow ? border0.getBottomLine() : 
						bottom02 != null ? bottom02 : border0.getHorizontalLine();
					return pickStyle(font0, fill0, left0, top0, right0, bottom0);
				}
				return pickStyle(font0, fill0, null, null, null, null);
			}
			return true;
		}

		private boolean pickStyle(SFont font0, SFill fill0, 
			SBorderLine left0, SBorderLine top0, SBorderLine right0, SBorderLine bottom0) {
			if (font == null)
				font = font0;
			if (fill == null)
				fill = fill0;
			if (left == null)
				left = left0;
			if (top == null)
				top = top0;
			if (right == null)
				right = right0;
			if (bottom == null)
				bottom = bottom0;
			return font != null && fill != null && left != null && top != null && right != null && bottom != null;
		}
		
		private SCellStyle getCellStyle() {
			final SBorder border0 = left == null && top == null 
					&& right == null && bottom == null? null : 
						new BorderImpl(left, top, right, bottom, null, null, null);
			return font == null && fill == null && border0 == null ? 
					null : new CellStyleImpl((AbstractFontAdv)font, (AbstractFillAdv)fill, (AbstractBorderAdv)border0);
		}
	}
	
	//ZSS-977
	@Override
	public SCellStyle getCellStyle(int row, int col) {
		final CellStylePicker picker = new CellStylePicker();
		final CellRegion all = _region.getRegion();
		final STableStyle tbStyle = _tableStyleInfo.getTableStyle();
		final boolean firstCol = col == all.getColumn();
		final boolean lastCol = col == all.getLastColumn();
		final boolean firstRow = row == all.getRow();
		final boolean lastRow = row == all.getLastRow();
		final boolean headerRow = firstRow && getHeaderRowCount() > 0;
		final boolean totalRow = lastRow && getTotalsRowCount() > 0;
		if (totalRow) {
			if (_tableStyleInfo.isShowLastColumn() && lastCol) {
				//Last Total Cell
				final STableStyleElem result = tbStyle.getLastTotalCellStyle();
				if (picker.pickDataStyle(result, null)) {
					return picker.getCellStyle();
				}
			} 
			if (_tableStyleInfo.isShowFirstColumn() && firstCol) { 
				//First Total Cell
				final STableStyleElem result = tbStyle.getFirstTotalCellStyle();
				if (picker.pickDataStyle(result, null)) {
					return picker.getCellStyle();
				}
			}
			//Total Row
			final STableStyleElem result = tbStyle.getTotalRowStyle();
			if (picker.pickDataStyle(result, null)) {
				return picker.getCellStyle();
			}
		} else if (headerRow) {
			if (_tableStyleInfo.isShowLastColumn() && lastCol) {
				//Last Header Cell
				final STableStyleElem result = _tableStyleInfo.getTableStyle().getLastHeaderCellStyle();
				if (picker.pickDataStyle(result, null)) {
					return picker.getCellStyle();
				}
			}
			if (_tableStyleInfo.isShowFirstColumn() && firstCol) {
				//First Header Cell
				final STableStyleElem result = tbStyle.getFirstHeaderCellStyle();
				if (picker.pickDataStyle(result, null)) {
					return picker.getCellStyle();
				}
			}
			//Header Row
			final STableStyleElem result = _tableStyleInfo.getTableStyle().getHeaderRowStyle();
			if (picker.pickDataStyle(result, null)) {
				return picker.getCellStyle();
			}
		}
		
		final boolean nextTotalsRow = getTotalsRowCount() > 0 && row == (all.getLastRow() - getTotalsRowCount()); 
		if (_tableStyleInfo.isShowFirstColumn() && firstCol) {
			//First Column
			final STableStyleElem result = _tableStyleInfo.getTableStyle().getFirstColumnStyle();
			if (picker.pickDataStyle(result, nextTotalsRow ? tbStyle.getTotalRowStyle() : null)) {
				return picker.getCellStyle();
			}
		}
		
		if (_tableStyleInfo.isShowLastColumn() && lastCol) {  
			//Last Column
			final STableStyleElem result = _tableStyleInfo.getTableStyle().getLastColumnStyle();
			if (picker.pickDataStyle(result, nextTotalsRow ? tbStyle.getTotalRowStyle() : null)) {
				return picker.getCellStyle();
			}
		}
		
		if (!headerRow && !totalRow) {
			//Row Stripe
			if (_tableStyleInfo.isShowRowStripes()) {
				final int topDataRow = all.getRow() + getHeaderRowCount();
				final STableStyle nmTableStyle = _tableStyleInfo.getTableStyle();
				final int rowStripe1Size = nmTableStyle.getRowStrip1Size();
				final int rowStripe2Size = nmTableStyle.getRowStrip2Size();
				int rowStripeSize = (row - topDataRow) % (rowStripe1Size + rowStripe2Size);
				final STableStyleElem result = 
						rowStripeSize < rowStripe1Size ?  // rowStripe1
							nmTableStyle.getRowStripe1Style():
							nmTableStyle.getRowStripe2Style();
				if (picker.pickDataStyle(result, nextTotalsRow ? tbStyle.getTotalRowStyle() : null)) {
					return picker.getCellStyle();
				}
			} 
			//Column Stripe
			if (_tableStyleInfo.isShowColumnStripes()) {
				final STableStyle nmTableStyle = _tableStyleInfo.getTableStyle();
				final int colStripe1Size = nmTableStyle.getColStrip1Size();
				final int colStripe2Size = nmTableStyle.getColStrip2Size();
				int colStripeSize = (col - all.getColumn()) % (colStripe1Size + colStripe2Size);
				final STableStyleElem result = 
						colStripeSize < colStripe1Size ? // colStripe1
							nmTableStyle.getColStripe1Style():
							nmTableStyle.getColStripe2Style();
				if (picker.pickDataStyle(result, nextTotalsRow ? tbStyle.getTotalRowStyle() : null)) {
					return picker.getCellStyle();
				}
			}
		}
		
		//Whole Table
		final STableStyleElem result = _tableStyleInfo.getTableStyle().getWholeTableStyle();
		picker.pickWholeTableStyle(result, 
				nextTotalsRow ? tbStyle.getTotalRowStyle() : null, 
						firstRow, firstCol, lastRow, lastCol);
		return picker.getCellStyle();
	}
}