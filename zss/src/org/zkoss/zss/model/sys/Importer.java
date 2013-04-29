/* Importer.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 13, 2007 2:42:02 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/

package org.zkoss.zss.model.sys;

import java.io.File;
import java.io.InputStream;

import org.zkoss.poi.ss.usermodel.Workbook;
/**
 * Importer class that used to import a input stream into a ZK Spreadsheet {@link Book}.
 * @author henrichen
 *
 */
public interface Importer {
	/**
	 * Imports a file into a spreadsheet book.
	 * @param filename a filename
	 * @return the {@link Workbook}
	 */
	public Book imports(String filename);
	/**
	 * Imports a file into a spreadsheet book.
	 * @param file a file
	 * @return the {@link Workbook}
	 */
	public Book imports(File file);
	/**
	 * Imports an input stream into a spreadsheet book.
	 * @param is inputstream
	 * @param bookname the book name
	 * @return the {@link Workbook}
	 */
	public Book imports(InputStream is, String bookname);
}
