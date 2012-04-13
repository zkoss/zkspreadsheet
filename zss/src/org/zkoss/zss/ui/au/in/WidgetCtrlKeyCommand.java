/* WidgetCtrlKeyCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 24, 2011 5:59:53 PM , Created by sam
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
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.KeyEvent;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.DrawingManager;
import org.zkoss.zss.model.impl.SheetCtrl;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author sam
 *
 */
public class WidgetCtrlKeyCommand implements Command {

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		final Map data = request.getData();
		if (data == null || data.size() != 6)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA,
				new Object[] {Objects.toString(data), this});
		
		String widgetType = (String) data.get("wgt");
		Worksheet sheet = ((Spreadsheet) comp).getSelectedSheet();
		if ("chart".equals(widgetType)) {
			processChart(sheet, data);
		} else if ("image".equals(widgetType)) {
			processPicture(sheet, data);
		}
	}
	
	private void processChart(Worksheet sheet, Map data) {
		String widgetId = (String) data.get("id");
		int keyCode = (Integer) data.get("keyCode");
		DrawingManager dm = ((SheetCtrl)sheet).getDrawingManager();
		if (KeyEvent.DELETE == keyCode) {
			List<Chart> charts = dm.getCharts();
			for (Chart chart : charts) {
				if (chart != null && chart.getChartId().equals(widgetId)) {
					Ranges.range(sheet).deleteChart(chart);
					break;
				}
			}
		}
	}
	
	private void processPicture(Worksheet sheet, Map data) {
		String widgetId = (String) data.get("id");
		int keyCode = (Integer) data.get("keyCode");
		DrawingManager dm = ((SheetCtrl)sheet).getDrawingManager();
		if (KeyEvent.DELETE == keyCode) {
			List<Picture> pics = dm.getPictures();
			for (Picture pic : pics) {
				if (pic.getPictureId().equals(widgetId)) {
					Ranges.range(sheet).deletePicture(pic);
					break;
				}
			}
		}
	}
}
