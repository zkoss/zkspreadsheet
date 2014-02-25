package org.zkoss.zss.ngapi.impl;

import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.zss.model.*;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SCellStyle.BorderType;
import org.zkoss.zss.model.SChart.NChartGrouping;
import org.zkoss.zss.model.SChart.NChartLegendPosition;
import org.zkoss.zss.model.SChart.NChartType;
import org.zkoss.zss.model.SHyperlink.HyperlinkType;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.ngapi.*;
import org.zkoss.zss.ngapi.NRange.SortDataOption;

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
	public SHyperlink getHyperlink() {
		
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
	public SSheet getSheet() {
		
		return null;
	}

	@Override
	public void setCellStyle(SCellStyle style) {
		

	}

	@Override
	public void fill(NRange dstRange, FillType fillType) {
		

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
	public SAutoFilter enableAutoFilter(int field, FilterOp filterOp,
			Object criteria1, Object criteria2, Boolean showButton) {
		
		return null;
	}

	@Override
	public SAutoFilter enableAutoFilter(boolean enable) {
		
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
	public SPicture addPicture(ViewAnchor anchor, byte[] image, Format format) {
		
		return null;
	}

	@Override
	public void deletePicture(SPicture picture) {
		

	}

	@Override
	public void movePicture(SPicture picture, ViewAnchor anchor) {
		

	}

	@Override
	public SDataValidation validate(String txt) {
		
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
	public SSheet createSheet(String name) {
		
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
	public SCellStyle getCellStyle() {
		
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
	
	public SChart addChart(ViewAnchor anchor, NChartType type, NChartGrouping grouping, NChartLegendPosition pos, boolean isThreeD) {
		return null;
	};
	
	@Override
	public void moveChart(SChart chart, ViewAnchor anchor) {
		
	}

	@Override
	public void deleteChart(SChart chart) {
		
	}

	@Override
	public void updateChart(SChart chart) {
		
	}

	@Override
	public void sort(NRange rng1, boolean desc1, NRange rng2, int type, boolean desc2,
			NRange rng3, boolean desc3, int header, int orderCustom, boolean matchCase,
			boolean sortByRows, int sortMethod, SortDataOption dataOption1, SortDataOption dataOption2,
			SortDataOption dataOption3) {
		
	}

}
