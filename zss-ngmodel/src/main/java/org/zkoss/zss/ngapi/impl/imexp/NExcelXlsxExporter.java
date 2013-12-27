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

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
/**
 * 
 * @author dennis, kuro
 * @since 3.5.0
 */
public class NExcelXlsxExporter extends AbstractExcelExporter {
	
	protected void exportColumnArray(NSheet sheet, Sheet poiSheet, NColumnArray columnArr) {
		XSSFSheet xssfSheet = (XSSFSheet) poiSheet;
		
        CTWorksheet ctSheet = xssfSheet.getCTWorksheet();
    	if(xssfSheet.getCTWorksheet().sizeOfColsArray() == 0) {
    		xssfSheet.getCTWorksheet().addNewCols();
    	}
    		
    	CTCol col = ctSheet.getColsArray(0).addNewCol();
        col.setMin(columnArr.getIndex()+1);
        col.setMax(columnArr.getLastIndex()+1);
    	col.setStyle(toPOICellStyle(columnArr.getCellStyle()).getIndex());
    	col.setCustomWidth(true);
    	col.setWidth(XUtils.pxToCTChar(columnArr.getWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
    	col.setHidden(columnArr.isHidden());
	}

	@Override
	protected Workbook createPoiBook() {
		return new XSSFWorkbook();
	}

}
