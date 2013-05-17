/* WidgetCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 18, 2011 5:20:37 PM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;

import java.util.List;
import java.util.Map;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.ZssChartX;
import org.zkoss.poi.xssf.usermodel.XSSFClientAnchor;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.DrawingManager;
import org.zkoss.zss.model.sys.impl.SheetCtrl;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author sam
 *
 */
public class WidgetUpdateCommand implements Command {

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		final Map data = request.getData();
		if (data == null || data.size() != 11)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA,
				new Object[] {Objects.toString(data), this});
		
		String type = (String) data.get("type");
		if ("onWidgetMove".equals(type) || "onWidgetSize".equals(type)) {
			processWidgetMove(((Spreadsheet) comp).getSelectedXSheet(), data);
		}
	}

	private void processWidgetMove(XSheet sheet, Map data) {
		String widgetType = (String) data.get("wgt");
		if ("chart".equals(widgetType)) {
			processChartMove(sheet, data);
		} else if ("image".equals(widgetType)) {
			processPictureMove(sheet, data);
		}
	}
	
	private void processChartMove(XSheet sheet, Map data) {
		DrawingManager dm = ((SheetCtrl)sheet).getDrawingManager();
		String widgetId = (String) data.get("id");
		ZssChartX chartX = null;
		List<ZssChartX> charts = dm.getChartXs();
		for (ZssChartX c : charts) {
			if (c != null) {
				if (c.getChartId().equals(widgetId)) {
					chartX = c;
					break;
				}	
			}
		}
		Chart chart = chartX.getChart();
		if (chart != null) {
			int dx1 = (Integer) data.get("dx1");
			int dy1 = (Integer) data.get("dy1");
			int dx2 = (Integer) data.get("dx2");
			int dy2 = (Integer) data.get("dy2");
			int col1 = (Integer) data.get("col1");
			int row1 = (Integer) data.get("row1");
			int col2 = (Integer) data.get("col2");
			int row2 = (Integer) data.get("row2");
			XRanges.range(sheet).moveChart(chart, 
					new XSSFClientAnchor(pxToEmu(dx1), pxToEmu(dy1), pxToEmu(dx2), pxToEmu(dy2), col1, row1, col2, row2));
		}
	}
	
	private void processPictureMove(XSheet sheet, Map data) {
		DrawingManager dm = ((SheetCtrl)sheet).getDrawingManager();
		String widgetId = (String) data.get("id");
		Picture pic = null;
		List<Picture> pics = dm.getPictures();
		for (Picture p : pics) {
			if (p.getPictureId().equals(widgetId)) {
				pic = p;
				break;
			}
		}
		if (pic != null) {
			int dx1 = (Integer) data.get("dx1");
			int dy1 = (Integer) data.get("dy1");
			int dx2 = (Integer) data.get("dx2");
			int dy2 = (Integer) data.get("dy2");
			int col1 = (Integer) data.get("col1");
			int row1 = (Integer) data.get("row1");
			int col2 = (Integer) data.get("col2");
			int row2 = (Integer) data.get("row2");
			XRanges.range(sheet).movePicture(pic, 
					new XSSFClientAnchor(pxToEmu(dx1), pxToEmu(dy1), pxToEmu(dx2), pxToEmu(dy2), col1, row1, col2, row2));
		}
	}
	
	public static int emuToPx(int emu) {
		return (int) Math.round(((double)emu) * 96 / 72 / 20 / 635); //assume 96dpi
	}
	
	/** convert pixel to EMU */
	public static int pxToEmu(int px) {
		return (int) Math.round(((double)px) * 72 * 20 * 635 / 96); //assume 96dpi
	}
}
