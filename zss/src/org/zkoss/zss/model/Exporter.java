/* BorderDrawEvent.java

	Purpose:
		
	Description:
		
	History:
		August 23, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model;

import java.io.OutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.zkoss.zss.model.Book;

import com.itextpdf.text.DocumentException;

/**
 * This interface defines methods for converting POI excel data model {@link Workbook} into another 
 * format written to a  {@link OutputStream}.
 * @author ashish
 *
 */
public interface Exporter {

	/**
	 * Converts Apache POI excel data model {@link Workbook} into another format written to a 
	 * {@link OutputStream}.
	 * @param workbook
	 * @param outputStream
	 * @throws DocumentException 
	 */
	public void export(Book workbook, OutputStream outputStream);
}
