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
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
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
	public NBook imports(InputStream is, String bookName) throws IOException {
		workbook = new XSSFWorkbook(is);
		book = NBooks.createBook(bookName);
		// import book scope content
		for(Sheet poiSheet : (XSSFWorkbook)workbook) { //only XSSFWorkwork is Iterable
			importPoiSheet(poiSheet);
		}
		
		importNameRange();
		return book;
	}

	/**
	 * get last column index that is configured, e.g. 16383
	 * reference BookHelper.getMaxConfiguredColumn()
	 * @param poiSheet
	 * @return
	 */
	@Override
	protected int getMaxConfiguredColumn(Sheet poiSheet) {
		int max = -1;
		CTWorksheet worksheet = ((XSSFSheet)poiSheet).getCTWorksheet();
		if(worksheet.sizeOfColsArray()<=0){
			return max;
		}
		CTCols colsArray = worksheet.getColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol colArray = colsArray.getColArray(i);
			max = Math.max(max, (int) colArray.getMax()-1);//in CT col it is 1base, -1 to 0base.
		}
		return max;
	}
	
}
 