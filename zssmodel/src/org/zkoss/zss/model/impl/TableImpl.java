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

	//ZSS-977
	private List<SCellStyle> getCellStylesInRange(int row, int col) {
		final List<SCellStyle> cellStyles = new ArrayList<SCellStyle>();
		final CellRegion all = _region.getRegion();
		final STableStyle tbStyle = _tableStyleInfo.getTableStyle();
		if (getTotalsRowCount() > 0 && row == all.getLastRow()) {
			if (_tableStyleInfo.isShowLastColumn() && col == all.getLastColumn()) {
				//Last Total Cell
				final STableStyleElem result = tbStyle.getLastTotalCellStyle();
				if (result != null) cellStyles.add(result);
			} 
			if (_tableStyleInfo.isShowFirstColumn() && col == all.getColumn()) { 
				//First Total Cell
				final STableStyleElem result = tbStyle.getFirstTotalCellStyle();
				if (result != null) cellStyles.add(result);
			}
		} 
		if (getHeaderRowCount() > 0 && row == all.getRow()) {
			if (_tableStyleInfo.isShowLastColumn() && col == all.getLastColumn()) {
				//Last Header Cell
				final STableStyleElem result = _tableStyleInfo.getTableStyle().getLastHeaderCellStyle();
				if (result != null) cellStyles.add(result);
			}
			if (_tableStyleInfo.isShowFirstColumn() && col == all.getColumn()) {
				//First Header Cell
				final STableStyleElem result = tbStyle.getFirstHeaderCellStyle();
				if (result != null) cellStyles.add(result);
			}
		}
		if (getTotalsRowCount() > 0 && row == all.getLastRow()) {
			//Total Row
			final STableStyleElem result = tbStyle.getTotalRowStyle();
			if (result != null) cellStyles.add(result);
		}
		if (getHeaderRowCount() > 0 && row == all.getRow()) {
			//Header Row
			final STableStyleElem result = _tableStyleInfo.getTableStyle().getHeaderRowStyle();
			if (result != null) cellStyles.add(result);
		}
		if (_tableStyleInfo.isShowFirstColumn() && col == all.getColumn()) {
			//First Column
			final STableStyleElem result = _tableStyleInfo.getTableStyle().getFirstColumnStyle();
			if (result != null) cellStyles.add(result);
		} 
		if (_tableStyleInfo.isShowLastColumn() && col == all.getLastColumn()) {  
			//Last Column
			final STableStyleElem result = _tableStyleInfo.getTableStyle().getLastColumnStyle();
			if (result != null) cellStyles.add(result);
		} 
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
			if (result != null) cellStyles.add(result);
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
			if (result != null) cellStyles.add(result);
		} 
		//Whole Table
		final STableStyleElem result = _tableStyleInfo.getTableStyle().getWholeTableStyle();
		if (result != null) cellStyles.add(result);
		return cellStyles;
	}
	
	//ZSS-977
	public SFont getFont(int row, int col) {
		List<SCellStyle> styles = getCellStylesInRange(row, col);
		for (SCellStyle style : styles) {
			final SFont font = style.getFont();
			if (font != null) return font;
		}
		return null;
	}
	
	//ZSS-977
	private SBorderLine getBottomLine(SBorder border, int row, int col) {
		final SBorderLine line = border.getBottomLine();
		if (line != null) return line;
		else { //try horizontal border
			final CellRegion region = _region.getRegion();
			if (row < region.getLastRow()) { // inside
				return border.getHorizontalLine();
			}
		}
		return null;
	}

	//ZSS-977
	private SBorderLine getTopLine(SBorder border, int row, int col) {
		final SBorderLine line = border.getTopLine();
		if (line != null) return line;
		else { //try horizontal border
			final CellRegion region = _region.getRegion();
			if (row > region.getRow()) { // inside
				return border.getHorizontalLine();
			}
		}
		return null;
	}

	//ZSS-977
	private SBorderLine getLeftLine(SBorder border, int row, int col) {
		final SBorderLine line = border.getLeftLine();
		if (line != null) return line;
		else { //try horizontal border
			final CellRegion region = _region.getRegion();
			if (col > region.getColumn()) { // inside
				return border.getVerticalLine();
			}
		}
		return null;
	}

	//ZSS-977
	private SBorderLine getRightLine(SBorder border, int row, int col) {
		final SBorderLine line = border.getRightLine();
		if (line != null) return line;
		else { //try horizontal border
			final CellRegion region = _region.getRegion();
			if (col < region.getLastColumn()) { // inside
				return border.getVerticalLine();
			}
		}
		return null;
	}

	//ZSS-977
	private SBorderLine getDiagonalLine(SBorder border, int row, int col) {
		return border.getDiagonalLine();
	}
	
	//ZSS-977
	public SCellStyle getCellStyle(int row, int col) {
		SFill fill = null;
		SFont font = null;
		SBorder border = null;
		SBorderLine bottom = null;
		SBorderLine top = null;
		SBorderLine left = null;
		SBorderLine right = null;
		SBorderLine diagonal = null;
		List<SCellStyle> styles = getCellStylesInRange(row, col);
		for (SCellStyle style : styles) {
			if (fill == null) {
				SFill fill0 = style.getFill();
				if (fill0 != null) fill = fill0;
			}
			if (font == null) {
				SFont font0 = style.getFont();
				if (font0 != null) font = font0;
			}
			if (border == null) {
				SBorder border0 = style.getBorder();
				if (border0 != null) {
					if (bottom == null) {
						SBorderLine bottom0 = getBottomLine(border0, row, col);
						if (bottom0 != null) bottom = bottom0;
					}
					if (top == null) {
						SBorderLine top0 = getTopLine(border0, row, col);
						if (top0 != null) top = top0;
					}
					if (left == null) {
						SBorderLine left0 = getLeftLine(border0, row, col);
						if (left0 != null) left = left0;
					}
					if (right == null) {
						SBorderLine right0 = getRightLine(border0, row, col);
						if (right0 != null) right = right0;
					}
//					if (diagonal == null) {
//						SBorderLine diagonal0 = getDiagonalLine(border0, row, col);
//						if (diagonal0 != null) diagonal = diagonal0;
//					}
					if (bottom != null && top != null && left != null && right != null /*&& diagonal != null*/) {
						border = new BorderImpl(left, top, right, bottom, diagonal, null, null);
					}
				}
			}
			if (fill != null && font != null && border != null) break;
		}
		if (border == null) {
			if (bottom != null || top != null || left != null || right != null || diagonal != null) {
				border = new BorderImpl(left, top, right, bottom, diagonal, null, null);
			}
		}

		return font == null && fill == null && border == null ? 
				null : new CellStyleImpl((AbstractFontAdv)font, (AbstractFillAdv)fill, (AbstractBorderAdv)border);
	}
}