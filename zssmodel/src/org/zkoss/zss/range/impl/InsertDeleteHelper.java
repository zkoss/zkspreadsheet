/* InsertDeleteHelper.java

	Purpose:
		
	Description:
		
	History:
		Feb 18, 2014 Created by Pao Wang

Copyright (C) 2014 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.range.impl;

import java.util.Iterator;

import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColumn;
import org.zkoss.zss.model.SColumnArray;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.impl.AbstractCellAdv;
import org.zkoss.zss.model.impl.AbstractColumnArrayAdv;
import org.zkoss.zss.model.impl.AbstractRowAdv;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRange.DeleteShift;
import org.zkoss.zss.range.SRange.InsertCopyOrigin;
import org.zkoss.zss.range.SRange.InsertShift;

/**
 * A helper to perform insert/delete row/column/cells.
 * @author Pao
 */
public class InsertDeleteHelper extends RangeHelperBase {

	public InsertDeleteHelper(SRange range) {
		super(range);
	}

	public void delete(DeleteShift shift) {
		// just process on the first sheet even this range over multiple sheets

		// insert row/column/cell
		if(isWholeRow()) { // ignore insert direction
			sheet.deleteRow(getRow(), getLastRow());

		} else if(isWholeColumn()) { // ignore insert direction
			sheet.deleteColumn(getColumn(), getLastColumn());

		} else if(shift != DeleteShift.DEFAULT) { // do nothing if "DEFAULT", it's according to XRange.delete() spec.
			sheet.deleteCell(getRow(), getColumn(), getLastRow(), getLastColumn(), shift == DeleteShift.LEFT);
		}
	}

	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin) {
		// just process on the first sheet even this range over multiple sheets

		// insert row/column/cell
		if(isWholeRow()) { // ignore insert direction
			sheet.insertRow(getRow(), getLastRow());
			// copy style/formal/size
			if(copyOrigin == InsertCopyOrigin.FORMAT_LEFT_ABOVE) {
				if(getRow() - 1 >= 0) {
					copyRowStyle(getRow() - 1, getRow(), getLastRow());
				}
			} else if(copyOrigin == InsertCopyOrigin.FORMAT_RIGHT_BELOW) {
				if(getLastRow() + 1 <= sheet.getBook().getMaxRowIndex()) {
					copyRowStyle(getLastRow() + 1, getRow(), getLastRow());
				}
			}
		} else if(isWholeColumn()) { // ignore insert direction
			sheet.insertColumn(getColumn(), getLastColumn());

			// copy style/formal/size
			if(copyOrigin == InsertCopyOrigin.FORMAT_LEFT_ABOVE) {
				if(getColumn() - 1 >= 0) {
					copyColumnStyle(getColumn() - 1, getColumn(), getLastColumn());
				}
			} else if(copyOrigin == InsertCopyOrigin.FORMAT_RIGHT_BELOW) {
				if(getLastColumn() + 1 <= sheet.getBook().getMaxColumnIndex()) {
					copyColumnStyle(getLastColumn() + 1, getColumn(), getLastColumn());
				}
			}

		} else if(shift != InsertShift.DEFAULT) { // do nothing if "DEFAULT", it's according to XRange.insert() spec.
			sheet.insertCell(getRow(), getColumn(), getLastRow(), getLastColumn(), shift == InsertShift.RIGHT);
			// TODO copy formal/style >>> in SheetImpl
		}
	}

	private void copyRowStyle(int srcRowIdx, int rowIdx, int lastRowIdx) {
		// copy row *local* style/height
		SRow srcRow = sheet.getRow(srcRowIdx);
		if(!srcRow.isNull()) {
			for(int r = rowIdx; r <= lastRowIdx; ++r) {
				SRow row = sheet.getRow(r);
				// style
				SCellStyle style = ((AbstractRowAdv)srcRow).getCellStyle(true);
				if(style != null) {
					row.setCellStyle(style);
				}
				// height
				if(srcRow.isCustomHeight()) { // according to Excel behavior
					row.setHeight(srcRow.getHeight());
					row.setCustomHeight(true);
				}
			}
		}

		// copy cell *local* style/format
		Iterator<SCell> cellsInRow = sheet.getCellIterator(srcRowIdx);
		while(cellsInRow.hasNext()) {
			SCell srcCell = cellsInRow.next();
			SCellStyle cellStyle = ((AbstractCellAdv)srcCell).getCellStyle(true);
			if(cellStyle != null) {
				for(int r = rowIdx; r <= lastRowIdx; ++r) {
					sheet.getCell(r, srcCell.getColumnIndex()).setCellStyle(cellStyle);
				}
			}
		}
	}

	private void copyColumnStyle(int srcColumnIdx, int columnIdx, int lastColumnIdx) {
		// copy column *local* style/height
		SColumnArray srcColumnArray = sheet.getColumnArray(srcColumnIdx);
		if(srcColumnArray!=null) {
			for(int c = columnIdx; c <= lastColumnIdx; ++c) {
				SColumn row = sheet.getColumn(c);
				// style
				SCellStyle style = ((AbstractColumnArrayAdv)srcColumnArray).getCellStyle(true);
				if(style != null) {
					row.setCellStyle(style);
				}
				// height
				if(srcColumnArray.isCustomWidth()) { // according to Excel behavior
					row.setWidth(srcColumnArray.getWidth());
					row.setCustomWidth(true);
				}
			}
		}

		// cell style/format
		Iterator<SRow> srcRows = sheet.getRowIterator();
		while(srcRows.hasNext()) {
			int r = srcRows.next().getIndex();
			SCell srcCell = sheet.getCell(r, srcColumnIdx);
			if(!srcCell.isNull()) {
				SCellStyle cellStyle = ((AbstractCellAdv)srcCell).getCellStyle(true);
				if(cellStyle != null) {
					for(int c = columnIdx; c <= lastColumnIdx; ++c) {
						sheet.getCell(r, c).setCellStyle(cellStyle);
					}
				}
			}
		}
	}
}
