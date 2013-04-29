/* ExcelImporter.java

	Purpose:
		
	Description:
		
	History:
		Dec 13, 2010 11:06:01 PM, Created by ashish

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys.impl;

import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XExporter;
import org.zkoss.zss.model.sys.XSheet;

/**
 * Exports {@link XBook} contents as an Excel file.
 * @author ashish
 *
 */
public class ExcelExporter implements XExporter {

	/**
	 * Exports {@link XBook} as an Excel file to given {@link OutputStream}
	 */
	public void export(XBook workbook, OutputStream outputStream) {
		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	public void export(XSheet worksheet, OutputStream outputStream) {
		throw new UnsupportedOperationException();
	}

	public void exportSelection(XSheet worksheet, AreaReference area,
			OutputStream outputStream) {
		throw new UnsupportedOperationException();
	}
}
