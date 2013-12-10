/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl;

import java.io.IOException;
import java.io.InputStream;

import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.impl.BookImpl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NExcelXlsImporter extends AbstractImporter{


	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		Workbook workbook = new HSSFWorkbook(is);
		NBook book = new BookImpl(bookName);
		
		return book;
	}

}
