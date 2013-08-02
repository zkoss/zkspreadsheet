/* Sheet.java

	Purpose:
		
	Description:
		
	History:
		Nov 24, 2010 2:51:25 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys;

import java.util.List;

import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.Drawing;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.PivotTable;

/**
 * ZK Spreadsheet sheet.
 * @author henrichen
 *
 */
public interface XSheet extends org.zkoss.poi.ss.usermodel.Sheet {
	/**
	 * Returns the associated ZK Spreadsheet {@link XBook} of this ZK Spreadsheet Sheet. 
	 * @return the associated ZK Spreadsheet {@link XBook} of this ZK Spreadsheet Sheet.
	 */
    public XBook getBook();
    
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
    
    /**
     * Returns validations in this ZK Spreadsheet sheet.
     * @return validations in this ZK Spreadsheet sheet.
     */
    public List<DataValidation> getDataValidations();
    
    /**
     * Returns 
     * @return
     */
    public List<PivotTable> getPivotTables();
    
	// ZSS-358: remove drawing part
	void removeDrawingPatriarch(Drawing drawing);

}
