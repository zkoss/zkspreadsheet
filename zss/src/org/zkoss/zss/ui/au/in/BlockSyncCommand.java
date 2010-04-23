/* BlockSyncCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		January 10, 2008 03:10:40 PM , Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under Lesser GPL Version 2.1 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;


import java.util.Map;

import org.zkoss.lang.Objects;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.SpreadsheetInCtrl;

/**
 * A Command (client to server) for synchronizing block side
 * @author Dennis.Chen
 *
 */
public class BlockSyncCommand implements Command{
	final static String Command = "onZSSSyncBlock";

	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED,
					BlockSyncCommand.class);

		final Map data = (Map) request.getData();
		if (data == null || data.size() != 17)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), BlockSyncCommand.class });

		Spreadsheet spreadsheet = ((Spreadsheet) comp);
		SpreadsheetInCtrl ctrl = ((SpreadsheetInCtrl) spreadsheet.getExtraCtrl());

		String sheetId = (String) data.get("sheetId");
		int dpWidth = (Integer) data.get("dpWidth");// pixel value of data panel width
		int dpHeight = (Integer) data.get("dpHeight");// pixel value of data panel height
		int viewWidth = (Integer) data.get("viewWidth");// pixel value of view width(scrollpanel.clientWidth)
		int viewHeight = (Integer) data.get("viewHeight");// pixel value of value height

		int blockLeft = (Integer) data.get("blockLeft");
		int blockTop = (Integer) data.get("blockTop");
		int blockRight = (Integer) data.get("blockRight");// + blockLeft - 1;
		int blockBottom = (Integer) data.get("blockBottom");// + blockTop - 1;;

		int fetchLeft = (Integer) data.get("fetchLeft");
		int fetchTop = (Integer) data.get("fetchTop");
		int fetchWidth = (Integer) data.get("fetchWidth");
		int fetchHeight = (Integer) data.get("fetchHeight");

		int rangeLeft = (Integer) data.get("rangeLeft");// visible range
		int rangeTop = (Integer) data.get("rangeTop");
		int rangeRight = (Integer) data.get("rangeRight");
		int rangeBottom = (Integer) data.get("rangeBottom");
		
		ctrl.setLoadedRect(blockLeft, blockTop, blockRight, blockBottom);	
	}

	public String getCommand() {
		return Command;
	}
}