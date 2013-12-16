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

import org.zkoss.poi.hssf.usermodel.*;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zss.ngmodel.*;
/**
 * 
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelXlsImporter extends AbstractExcelImporter{


	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		workbook = new HSSFWorkbook(is);
		book = NBooks.createBook(bookName);
		for (int i = 0 ; i < workbook.getNumberOfSheets(); i++){
			importPoiSheet(workbook.getSheetAt(i));
		}
		return book;
	}

	@Override
	protected int getMaxConfiguredColumn(Sheet poiSheet) {
		// FIXME cannot impelement getMaxConfiguredColumn() because HSSFSheet.getSheet() is protected.
//		((HSSFSheetImpl)poiSheet).get
//		((HSSFSheet)poiSheet).getSheet().getMaxConfiguredColumn();
		return 100;
	}


}
