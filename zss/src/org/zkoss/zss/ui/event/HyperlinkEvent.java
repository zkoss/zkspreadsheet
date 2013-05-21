/* HyperlinkEvent.java

	Purpose:
		
	Description:
		
	History:
		Jul 19, 2010 2:08:02 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.ui.event;

import java.util.Map;

import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.au.AuRequests;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.XUtils;

/**
 * Event when end user click on the hyperlink of a cell (used with onHyperlink event).
 * @author henrichen
 */
public class HyperlinkEvent extends CellMouseEvent{
	private String _address;
	private HyperlinkType _type;

	public static HyperlinkEvent getHyperlinkEvent(AuRequest request) {
		final Map data = request.getData();
		final Component comp = request.getComponent();
		String sheetId = (String) data.get("sheetId");
		Sheet sheet = ((Spreadsheet) comp).getSelectedSheet();
		if (!XUtils.getSheetUuid(((SheetImpl)sheet).getNative()).equals(sheetId))
			return null;
		
		final String name = request.getCommand();
		final int keys = AuRequests.parseKeys(data);
		final int type = AuRequests.getInt(data, "type", 0, true);
		return new HyperlinkEvent(name, comp, sheet,
				AuRequests.getInt(data, "row", 0, true),
				AuRequests.getInt(data, "col", 0, true),
				(String) data.get("href"),
				EnumUtil.toHyperlinkType(type),
				AuRequests.getInt(data, "x", 0, true),
				AuRequests.getInt(data, "y", 0, true),
				AuRequests.getInt(data, "pageX", 0, true),
				AuRequests.getInt(data, "pageY", 0, true), keys);
	}
	public HyperlinkEvent(String name, Component target, Sheet sheet, int row ,int col, String address, HyperlinkType type, int x, int y,
			int pageX, int pageY, int keys) {
		super(name, target, x, y, keys, sheet, row, col, pageX, pageY);
		this._address = address;
		this._type = type;
	}

	/**
	 * LINK Reference.
	 * @return URI reference.
	 */
	public String getAddress() {
		return _address;
	}

	/**
	 * URL LINK type 
	 * @return link type
	 */
	public HyperlinkType getType() {
		return _type;
	}
}
