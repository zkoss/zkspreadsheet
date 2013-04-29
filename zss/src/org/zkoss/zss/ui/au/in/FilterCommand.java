/* FilterCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 5, 2011 9:44:00 AM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;

import java.util.Map;

import org.zkoss.json.JSONArray;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author sam
 *
 */
public class FilterCommand implements Command {

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		
		final Map data = request.getData();
		String type = (String) data.get("type");
		if ("onApplyFilter".equals(type)) {
			applyFilter(((Spreadsheet) comp).getSelectedSheet(), data);
		}
	}
	
	private void applyFilter (XSheet worksheet, Map data) {
		final boolean selectAll = (Boolean) data.get("all");
		final String cellRangeAddr = (String) data.get("range");
		final int field = (Integer) data.get("field");
		final XRange range = XRanges.range(worksheet, cellRangeAddr);
		
		if (selectAll) {
			range.autoFilter(field, null, AutoFilter.FILTEROP_VALUES, null, null);
		} else { //partial selection
			JSONArray ary = (JSONArray) data.get("criteria");
			range.autoFilter(field, ary.toArray(new String[ary.size()]), AutoFilter.FILTEROP_VALUES, null, null);
		}
	}
}
