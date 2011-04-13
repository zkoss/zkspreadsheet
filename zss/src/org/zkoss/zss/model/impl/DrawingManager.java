/**
 * 
 */
package org.zkoss.zss.model.impl;

import java.util.List;

import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.Combo;
import org.zkoss.poi.ss.usermodel.Picture;

/**
 * ZK Spreadsheet Drawing manager. 
 * @author henrichen
 *
 */
public interface DrawingManager {
	/**
	 * Return pictures in sheet.
	 * @return pictures in sheet.
	 */
	public List<Picture> getPictures();
	
	/**
	 * Returns charts in sheet.
	 * @return charts in sheet.
	 */
	public List<Chart> getCharts();

	public List<Combo> getCombos();
}
