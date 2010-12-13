/* ExcelImporter.java

	Purpose:
		
	Description:
		
	History:
		Dec 13, 2010 11:06:01 PM, Created by ashish

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/package org.zkoss.zss.model.impl;

import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Exporter;

/**
 * Exports {@link Book} contents as an Excel file.
 * @author ashish
 *
 */
public class ExcelExporter implements Exporter {

	/**
	 * Exports {@link Book} as an Excel file to given {@link OutputStream}
	 */
	public void export(Book workbook, OutputStream outputStream) {
		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	public void export(Sheet worksheet, OutputStream outputStream) {
		throw new UnsupportedOperationException();
	}

	public void exportSelection(Sheet worksheet, AreaReference area,
			OutputStream outputStream) {
		throw new UnsupportedOperationException();
	}
}
