/* HyperlinkHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 30, 2010 8:00:42 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.event;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

/**
 * A Helper class for Hyperlink
 * @author Sam
 *
 */
public class HyperlinkHelper {
	private HyperlinkHelper(){}
	
	public static void onHyperlink(ForwardEvent evt) {
		String param = (String)evt.getData();
		MainWindowCtrl ctrl = MainWindowCtrl.getInstance();
		Spreadsheet ss = ctrl.getSpreadsheet();
		if (param == null || ss == null) {
			return;
		}
		
		if (param.equals(Labels.getLabel("hyperlink"))) {
			HashMap arg = new HashMap();
			arg.put("spreadsheet", ss);
			Executions.createComponents("/menus/hyperlink/insertHyperlink.zul", ctrl.getMainWindow(), arg);
		} else if (param.equals(Labels.getLabel("hyperlink.remove"))){
			//TODO remove hyperlink
			
		}
	}
}
