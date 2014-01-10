/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.*;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zss.ngmodel.*;
/**
 * 
 * @author dennis, kuro
 * @since 3.5.0
 */
public class NExcelXlsExporter extends AbstractExcelExporter {
	
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.readLock().lock();

		try {
			workbook = new HSSFWorkbook();
			
			for(NSheet sheet : book.getSheets()) {
				exportSheet(sheet);
			}
			
			exportNamedRange(book);
			
			workbook.write(fos);
		} finally {
			lock.readLock().unlock();
		}
	}

	@Override
	protected void exportColumnArray(NSheet sheet, Sheet poiSheet, NColumnArray columnArr) {
		
		CellStyle poiCellStyle = toPOICellStyle(columnArr.getCellStyle());
		boolean hidden = columnArr.isHidden();
		
		for(int i = columnArr.getIndex(); i <= columnArr.getLastIndex() && i <= SpreadsheetVersion.EXCEL97.getMaxColumns(); i++) {
			poiSheet.setColumnWidth(i+1, XUtils.pxToFileChar256(columnArr.getWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
			poiSheet.setColumnHidden(i+1, hidden);
			poiSheet.setDefaultColumnStyle(i+1, poiCellStyle);
		}
	}

	@Override
	protected Workbook createPoiBook() {
		return new HSSFWorkbook();
	}

	@Override
	protected void exportChart(NSheet sheet, Sheet poiSheet) {
		// not support in XLS
	}
	
	@Override
	protected void exportPicture(NSheet sheet, Sheet poiSheet) {
		// not support in XLS
	}

	@Override
	protected void exportValidation(NSheet sheet, Sheet poiSheet) {
		// not support in XLS
	}

}
