/* EmptyRange.java

	Purpose:
		
	Description:
		
	History:
		Oct 28, 2010 5:48:49 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys.impl;

import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.poi.ss.usermodel.charts.LegendPosition;
import org.zkoss.zss.model.sys.XAreas;
import org.zkoss.zss.model.sys.XFormatText;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XSheet;

/**
 * Class to represent an empty Range.
 * @author henrichen
 */
public class EmptyRange implements XRange {

	@Override
	public void autoFill(XRange dstRange, int fillType) {
	}

	@Override
	public void borderAround(BorderStyle lineStyle, String color) {
	}

	@Override
	public void clearContents() {
	}

	@Override
	public XRange copy(XRange dstRange) {
		return null;
	}

	@Override
	public void delete(int shift) {
	}

	@Override
	public void fillDown() {
	}

	@Override
	public void fillLeft() {
	}

	@Override
	public void fillRight() {
	}

	@Override
	public void fillUp() {
	}

	@Override
	public XAreas getAreas() {
		return new AreasImpl();
	}

	@Override
	public XRange getCells(int row, int col) {
		return this;
	}

	@Override
	public int getColumn() {
		return -1;
	}

	@Override
	public XRange getColumns() {
		return this;
	}

	@Override
	public long getCount() {
		return 0L;
	}

	@Override
	public XRange getDependents() {
		return this;
	}

	@Override
	public XRange getDirectDependents() {
		return this;
	}

	@Override
	public String getEditText() {
		return null;
	}

	@Override
	public XSheet getSheet() {
		return null;
	}

	@Override
	public XFormatText getFormatText() {
		return null;
	}

	@Override
	public Hyperlink getHyperlink() {
		return null;
	}

	@Override
	public RichTextString getRichEditText() {
		return null;
	}

	@Override
	public int getRow() {
		return -1;
	}

	@Override
	public XRange getRows() {
		return this;
	}

	@Override
	public RichTextString getText() {
		return null;
	}

	@Override
	public void insert(int shift, int copyOrigin) {
	}

	@Override
	public void merge(boolean across) {
	}

	@Override
	public void move(int nRow, int nCol) {
	}

	@Override
	public XRange pasteSpecial(int pasteType, int operation, boolean SkipBlanks,
			boolean transpose) {
		return null;
	}

	@Override
	public XRange pasteSpecial(XRange dstRange, int pasteType, int pasteOp,
			boolean skipBlanks, boolean transpose) {
		return null;
	}

	@Override
	public void setBorders(short borderIndex, BorderStyle lineStyle,
			String color) {
	}

	@Override
	public void setColumnWidth(int char256) {
	}

	@Override
	public void setDisplayGridlines(boolean show) {
	}
	
	@Override
	public void protectSheet(String password) {
	}

	@Override
	public void setEditText(String txt) {
	}

	@Override
	public void setHidden(boolean hidden) {
	}

	@Override
	public void setHyperlink(int linkType, String address, String display) {
	}

	@Override
	public void setRichEditText(RichTextString txt) {
	}

	@Override
	public void setRowHeight(int points) {
	}

	@Override
	public void setStyle(CellStyle style) {
	}

	@Override
	public void sort(XRange rng1, boolean desc1, XRange rng2, int type,
			boolean desc2, XRange rng3, boolean desc3, int header,
			int orderCustom, boolean matchCase, boolean sortByRows,
			int sortMethod, int dataOption1, int dataOption2, int dataOption3) {
	}

	@Override
	public void unMerge() {
	}

	@Override
	public XRange getDirectPrecedents() {
		return this;
	}

	@Override
	public XRange getPrecedents() {
		return this;
	}

	@Override
	public int getLastColumn() {
		return -1;
	}

	@Override
	public int getLastRow() {
		return -1;
	}

	@Override
	public Object getValue() {
		return null;
	}

	@Override
	public void setValue(Object value) {
	}

	@Override
	public XRange getOffset(int rowOffset, int colOffset) {
		return this;
	}

	@Override
	public AutoFilter autoFilter() {
		return null;
	}

	@Override
	public AutoFilter autoFilter(int field, Object criteria1, int filterOp, Object criteria2, Boolean visibleDropDown) {
		return null;
	}

	@Override
	public XRange getCurrentRegion() {
		return this;
	}

	@Override
	public void applyFilter() {
	}

	@Override
	public void showAllData() {
	}

	@Override
	public Chart addChart(ClientAnchor anchor, ChartData data, ChartType type,
			ChartGrouping grouping, LegendPosition pos) {
		return null;
	}

	@Override
	public Picture addPicture(ClientAnchor anchor, byte[] image, int format) {
		return null;
	}

	@Override
	public void deletePicture(Picture picture) {
	}

	@Override
	public void movePicture(Picture picture, ClientAnchor anchor) {
	}

	@Override
	public void moveChart(Chart chart, ClientAnchor anchor) {
	}

	@Override
	public void deleteChart(Chart chart) {
	}

	@Override
	public DataValidation validate(String txt) {
		return null;
	}

	@Override
	public boolean isAnyCellProtected() {
		return true;
	}

	@Override
	public void notifyMoveFriendFocus(Object token) {
	}

	@Override
	public void notifyDeleteFriendFocus(Object token) {
	}

	@Override
	public void deleteSheet() {
	}

	@Override
	public void setRowHeight(int points, boolean customHeight) {
	}

	@Override
	public boolean isCustomHeight() {
		return false;
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.model.Range#createSheet(java.lang.String)
	 */
	@Override
	public void createSheet(String name) {
		// TODO Auto-generated method stub
		
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.model.Range#setSheetName(java.lang.String)
	 */
	@Override
	public void setSheetName(String name) {
		// TODO Auto-generated method stub
		
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.model.Range#setSheetOrder(int)
	 */
	@Override
	public void setSheetOrder(int pos) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean isWholeRow() {
		return false;
	}

	@Override
	public boolean isWholeColumn() {
		return false;
	}

	@Override
	public boolean isWholeSheet() {
		return false;
	}

	@Override
	public XRange findAutoFilterRange() {
		return null;
	}
}
