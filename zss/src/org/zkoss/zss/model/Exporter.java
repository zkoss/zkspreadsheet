/* BorderDrawEvent.java

	Purpose:
		
	Description:
		
	History:
		August 23, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.model;

import java.io.OutputStream;

import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.zss.model.sys.Book;
import org.zkoss.zss.model.sys.Worksheet;

/**
 * This interface defines methods for exporting ZK Spreadsheet {@link Book} into another 
 * format written to a  {@link OutputStream}.
 * @author ashish
 * @author dennischen
 * @deprecated since 3.0.0, please use class in package {@code org.zkoss.zss.api}
 */
public interface Exporter {

	/**
	 * Exports ZK Spreadsheet {@link Book} into another format written to a 
	 * {@link OutputStream}. Note that it exports entire workbook. 
	 * @param workbook
	 * @param outputStream
	 */
	public void export(Book workbook, OutputStream outputStream);

	/**
	 * Exports ZK Spreadsheet sheet into another format written to a 
	 * @param worksheet sheet instance that contains selected area
	 * @param outputStream outoutStream to which exported contents to be written
	 */
	public void export(Worksheet worksheet, OutputStream outputStream);

	/**
	 * Exports selected area of ZK Spreadsheet active sheet represented by 
	 * {@link org.zkoss.zss.model.sys.Range}
	 * @param worksheet sheet instance that contains selected area
	 * @param area area representing selected area to be exported
	 * @param outputStream outoutStream to which exported contents to be written
	 */
	public void exportSelection(Worksheet worksheet, AreaReference area, OutputStream outputStream);
}
