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
package org.zkoss.zss.ngapi;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.zss.ngmodel.NBook;

/**
 * An exporter to export a book to a out stream or file
 * @author dennis
 * @since 3.5.0
 */
public interface NExporter {
	/**
	 * Export book
	 * @param book the book to export
	 * @param fos the output stream to store data
	 * @throws IOException
	 */
	public void export(NBook book, OutputStream fos) throws IOException;
	
	/**
	 * Export book
	 * @param book the book to export
	 * @param fos the output file to store data
	 * @throws IOException
	 */
	public void export(NBook book, File file) throws IOException;
}
