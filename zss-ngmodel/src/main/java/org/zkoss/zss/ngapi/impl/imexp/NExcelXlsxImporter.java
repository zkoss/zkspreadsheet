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

import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.util.SpreadsheetVersion;
/**
 * Convert Excel XLSX format file to a Spreadsheet {@link NBook} model including following information:
 * Book:
 * 		name
 * Sheet:
 * 		name, column width, row height, hidden row (column), row (column) style
 * Cell:
 * 		type, value, font with color and style, type offset(normal or subscript), background color, border
 * 
 * TODO use XLSX, XLS common interface (e.g. CellStyle instead of {@link XSSFCellStyle}) to get content first for that codes can be easily moved to parent class. 
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelXlsxImporter extends AbstractExcelImporter{

	@Override
	protected Workbook createPoiBook(InputStream is) throws IOException {
		return new XSSFWorkbook(is);
	}

	/**
	 * get last column index that is different from default value, e.g. width is customized.
	 * reference BookHelper.getMaxConfiguredColumn()
	 * @param poiSheet
	 * @return
	 */
	protected int getLastChangedColumnIndex(Sheet poiSheet) {
		int max = -1;
		CTWorksheet worksheet = ((XSSFSheet)poiSheet).getCTWorksheet();
		if(worksheet.sizeOfColsArray()<=0){
			return max;
		}
		
		CTCols colsArray = worksheet.getColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol ctCol = colsArray.getColArray(i);
			//TODO ZSS-525, to avoid getting upper bound value of column index
			if (ctCol.getCustomWidth() && ctCol.getMax()!=SpreadsheetVersion.EXCEL2007.getMaxColumns()){
				max = Math.max(max, (int) ctCol.getMax()-1);//in CT col it is 1base, -1 to 0base.
			}
		}
		return max;
	}

	/**
	 * [ISO/IEC 29500-1 1st Edition] 18.3.1.13 col (Column Width & Formatting)
	 * By experiments, CT_Col is always created in ascending order by min and the range specified by min & max doesn't overlap each other.<br/>
	 * For example:
	 * <x:cols xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	 *   <x:col min="1" max="1" width="11.28515625" customWidth="1" />
	 *   <x:col min="2" max="4" width="11.28515625" hidden="1" customWidth="1" />
	 *   <x:col min="5" max="5" width="11.28515625" customWidth="1" />
	 * </x:cols> 
	 */
	@Override
	protected void importColumn(Sheet poiSheet, NSheet sheet, int defaultWidth) {
		CTWorksheet worksheet = ((XSSFSheet)poiSheet).getCTWorksheet();
		if(worksheet.sizeOfColsArray()<=0){
			return;
		}
		
		CTCols colsArray = worksheet.getColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol ctCol = colsArray.getColArray(i);
			//max is 16384
			
			NColumnArray columnArray = sheet.setupColumnArray((int)ctCol.getMin()-1, (int)ctCol.getMax()-1);
			boolean hidden = ctCol.getHidden();
			int columnIndex = (int)ctCol.getMin()-1;
			
			columnArray.setHidden(hidden);
			if (hidden == false){
				//when CT_Col is hidden with default width, We don't import the width for it's 0.  
				columnArray.setWidth(XUtils.getWidthAny(poiSheet, columnIndex, CHRACTER_WIDTH));
			}
			CellStyle columnStyle = poiSheet.getColumnStyle(columnIndex);
			if (columnStyle != null){
				columnArray.setCellStyle(importCellStyle(columnStyle));
			}
		}
	}
	
}
 