/**
 * 
 */
package org.zkoss.zss.model.impl;

import java.util.List;

import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.ZssChartX;
import org.zkoss.poi.ss.usermodel.Combo;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.poi.ss.usermodel.charts.LegendPosition;
import org.zkoss.zss.model.Worksheet;

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
	public List<ZssChartX> getCharts();

	public List<Combo> getCombos();
	
	public ZssChartX addChartX(Worksheet sheet, ClientAnchor anchor, ChartData data, ChartType type,
			ChartGrouping grouping, LegendPosition pos);
	
	public Picture addPicture(Worksheet sheet, ClientAnchor anchor, byte[] imageData, int format);
	
	public void deletePicture(Worksheet sheet, Picture picture);
}
