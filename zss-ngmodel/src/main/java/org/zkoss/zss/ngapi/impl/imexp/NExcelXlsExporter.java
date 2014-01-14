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
	protected void exportColumnArray(NSheet sheet, Sheet poiSheet, NColumnArray columnArr) {
		
		CellStyle poiCellStyle = toPOICellStyle(columnArr.getCellStyle());
		boolean hidden = columnArr.isHidden();
		
		for(int i = columnArr.getIndex(); i <= columnArr.getLastIndex() && i <= SpreadsheetVersion.EXCEL97.getMaxColumns(); i++) {
			poiSheet.setColumnWidth(i, UnitUtil.pxToFileChar256(columnArr.getWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
			poiSheet.setColumnHidden(i, hidden);
			poiSheet.setDefaultColumnStyle(i, poiCellStyle);
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
