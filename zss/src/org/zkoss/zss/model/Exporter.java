/* BorderDrawEvent.java

	Purpose:
		
	Description:
		
	History:
		August 23, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model;

import java.io.OutputStream;
import org.zkoss.zss.model.Book;

/**
 * This interface defines methods for exporting ZK Spreadsheet {@link Book} into another 
 * format written to a  {@link OutputStream}.
 * @author ashish
 *
 */
public interface Exporter {

	/**
	 * Exports ZK Spreadsheet {@link Book} into another format written to a 
	 * {@link OutputStream}.
	 * @param workbook
	 * @param outputStream
	 * @throws DocumentException 
	 */
	public void export(Book workbook, OutputStream outputStream);
}
