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

import java.util.Map;

import org.zkoss.lang.Objects;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.KeyEvent;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.WidgetKeyEvent;
import org.zkoss.zss.ui.impl.XUtils;

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
		if (data == null || data.size() != 7)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA,
				new Object[] {Objects.toString(data), this});
		
		
		Sheet sheet = ((Spreadsheet) comp).getSelectedSheet();
		
		String sheetId= (String) data.get("sheetId");
		if (!XUtils.getSheetUuid(sheet).equals(sheetId))
			return;
		
		String widgetType = (String) data.get("wgt");
		org.zkoss.zk.ui.event.KeyEvent evt = KeyEvent.getKeyEvent(request);
		Object widgetData = null;
		
		String id = (String) data.get("id");
		
		if ("image".equals(widgetType)) {
			for(Picture p:Ranges.range(sheet).getPictures()){
				if(p.getId().equals(id)){
					widgetData = p;
					break;
				}
			}
			
		} else if ("chart".equals(widgetType)) {
			for(Chart c:Ranges.range(sheet).getCharts()){
				if(c.getId().equals(id)){
					widgetData = c;
					break;
				}
			}
		}
		if(widgetData==null){
			//TODO ignore it or throw exception?
			return;
		}
		Event zssKeyEvt = new WidgetKeyEvent(evt.getName(), evt.getTarget(), sheet,widgetData,
				evt.getKeyCode(), evt.isCtrlKey(), evt.isShiftKey(), evt.isAltKey());
		Events.postEvent(zssKeyEvt);
		
	}
}
