package org.zkoss.zss.ngapi.impl;

import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NChart.NChartGrouping;
import org.zkoss.zss.ngmodel.NChart.NChartLegendPosition;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.NHyperlink.HyperlinkType;
import org.zkoss.zss.ngmodel.NPicture.Format;
import org.zkoss.zss.ngmodel.chart.NChartData;

/**
 * the empty range implementation that do nothing
 * @author Dennis
 * @since 3.5.0
 */
/*package*/ class NEmptyRange implements NRange {

	@Override
	public ReadWriteLock getLock() {
		
		return null;
	}

	@Override
	public NHyperlink getHyperlink() {
		
		return null;
	}

	@Override
	public String getEditText() {
		
		return null;
	}

	@Override
	public void setEditText(String txt) {
		

	}

	@Override
	public NRange copy(NRange dstRange, boolean cut) {
		
		return null;
	}

	@Override
	public NRange copy(NRange dstRange) {
		
		return null;
	}

	@Override
	public NRange pasteSpecial(NRange dstRange, PasteType pasteType,
			PasteOperation pasteOp, boolean skipBlanks, boolean transpose) {
		
		return null;
	}

	@Override
	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin) {
		

	}

	@Override
	public void delete(DeleteShift shift) {
		

	}

	@Override
	public void merge(boolean across) {
		

	}

	@Override
	public void unmerge() {
		

	}

	@Override
	public void setBorders(ApplyBorderType borderIndex, BorderType lineStyle,
			String color) {
		

	}

	@Override
	public void move(int nRow, int nCol) {
		

	}

	@Override
	public void setColumnWidth(int widthPx) {
		

	}

	@Override
	public void setRowHeight(int heightPx) {
		

	}

	@Override
	public void setColumnWidth(int widthPx, boolean custom) {
		

	}

	@Override
	public void setRowHeight(int heightPx, boolean custom) {
		

	}

	@Override
	public NSheet getSheet() {
		
		return null;
	}

	@Override
	public void setCellStyle(NCellStyle style) {
		

	}

	@Override
	public void autoFill(NRange dstRange, AutoFillType fillType) {
		

	}

	@Override
	public void clearContents() {
		

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
	public NRange findAutoFilterRange() {
		
		return null;
	}

	@Override
	public NAutoFilter enableAutoFilter(int field, FilterOp filterOp,
			Object criteria1, Object criteria2, Boolean showButton) {
		
		return null;
	}

	@Override
	public NAutoFilter enableAutoFilter(boolean enable) {
		
		return null;
	}

	@Override
	public void resetAutoFilter() {
		

	}

	@Override
	public void applyAutoFilter() {
		

	}

	@Override
	public void setHidden(boolean hidden) {
		

	}

	@Override
	public void setDisplayGridlines(boolean show) {
		

	}

	@Override
	public void protectSheet(String password) {
		

	}

	@Override
	public void setHyperlink(HyperlinkType linkType, String address,
			String display) {
		

	}

	@Override
	public NRange getColumns() {
		
		return this;
	}

	@Override
	public NRange getRows() {
		
		return this;
	}

	@Override
	public int getRow() {
		
		return -1;
	}

	@Override
	public int getColumn() {
		
		return -1;
	}

	@Override
	public int getLastRow() {
		
		return -1;
	}

	@Override
	public int getLastColumn() {
		
		return -1;
	}

	@Override
	public void setValue(Object value) {
		

	}

	@Override
	public Object getValue() {
		
		return null;
	}

	@Override
	public NRange getOffset(int rowOffset, int colOffset) {
		
		return this;
	}

	@Override
	public NPicture addPicture(NViewAnchor anchor, byte[] image, Format format) {
		
		return null;
	}

	@Override
	public void deletePicture(NPicture picture) {
		

	}

	@Override
	public void movePicture(NPicture picture, NViewAnchor anchor) {
		

	}

	@Override
	public NDataValidation validate(String txt) {
		
		return null;
	}

	@Override
	public boolean isAnyCellProtected() {
		
		return false;
	}

	@Override
	public void notifyCustomEvent(String customEventName, Object data,
			boolean writeLock) {
		

	}

	@Override
	public void deleteSheet() {
		

	}

	@Override
	public NSheet createSheet(String name) {
		
		return null;
	}

	@Override
	public void setSheetName(String name) {
		

	}

	@Override
	public void setSheetOrder(int pos) {
		

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
	public void notifyChange() {
		

	}

	@Override
	public void setFreezePanel(int numOfRow, int numOfColumn) {
		

	}

	@Override
	public String getCellFormatText() {
		
		return null;
	}

	@Override
	public NCellStyle getCellStyle() {
		
		return null;
	}

	@Override
	public boolean isSheetProtected() {
		
		return false;
	}

	@Override
	public String getCellDataFormat() {
		return null;
	}
	
	public NChart addChart(NViewAnchor anchor, NChartData data, NChartType type, NChartGrouping grouping, NChartLegendPosition pos) {
		return null;
	};
	
	@Override
	public void moveChart(NChart chart, NViewAnchor anchor) {
		
	}

	@Override
	public void deleteChart(NChart chart) {
		
	}

}
