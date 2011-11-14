/* Sheet.java

	Purpose:
		
	Description:
		
	History:
		Nov 24, 2010 2:51:25 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;

import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.Picture;


/**
 * ZK Spreadsheet sheet.
 * @author henrichen
 *
 */
public interface Worksheet extends org.zkoss.poi.ss.usermodel.Sheet {
	/**
	 * Returns the associated ZK Spreadsheet {@link Book} of this ZK Spreadsheet Sheet. 
	 * @return the associated ZK Spreadsheet {@link Book} of this ZK Spreadsheet Sheet.
	 */
    public Book getBook();
    
    /**
     * Returns pictures in this ZK Spreadsheet sheet.
     * @return pictures in this ZK Spreadsheet sheet.
     */
    public List<Picture> getPictures();
    
    /**
     * Returns charts in this ZK Spreadsheet sheet.
     * @return charts in this ZK Spreadsheet sheet.
     */
    public List<Chart> getCharts();
}
