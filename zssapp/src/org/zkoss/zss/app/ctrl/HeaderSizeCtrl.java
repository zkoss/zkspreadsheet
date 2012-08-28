/* HeaderSizeCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 20, 2010 3:39:33 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ctrl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zul.Button;
import org.zkoss.zul.Intbox;
import org.zkoss.zul.Label;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class HeaderSizeCtrl extends GenericForwardComposer {
	
	Dialog _headerSizeDialog;
	/* header type to set, specify column or row */
	private final static String KEY_HEADER_TYPE = "org.zkoss.zss.app.ctrl.HeaderSizeCtrl.headerType";
	private Integer headerType;
	
	/* header original value */
	private final static String KEY_HEADER_SIZE = "org.zkoss.zss.app.ctrl.HeaderSizeCtrl.headerSize";
	private static final String KEY_SELECTION = "org.zkoss.zss.app.ctrl.HeaderSizeCtrl.selection";
	private Integer size;
	
	/* header title specify Column width or Row height */
	private Label title;
	
	/* header size of the target */
	private Intbox headerSize;
	private Rect selection;
	
	private Button okBtn;
	
	/**
	 * Construct arguments for {@link onOpen} event data
	 * @param headrType
	 * @param headerSize
	 * @return
	 */
	public static Map newArg(Integer headrType, Integer headerSize, Rect selection) {
		Map arg = new HashMap();
		arg.put(HeaderSizeCtrl.KEY_HEADER_TYPE, headrType);
		arg.put(HeaderSizeCtrl.KEY_HEADER_SIZE, headerSize);
		arg.put(HeaderSizeCtrl.KEY_SELECTION, selection);
		return arg;
	}
	
	public void onOpen$_headerSizeDialog(ForwardEvent event) {
		Map arg = (Map) event.getOrigin().getData();
		selection = (Rect)arg.get(KEY_SELECTION);
		headerType = (Integer)arg.get(KEY_HEADER_TYPE);
		initTitle(headerType);
		size = (Integer)arg.get(KEY_HEADER_SIZE);
		headerSize.setValue(size);
		
		_headerSizeDialog.setMode(Window.MODAL);
	}
	
	private void initTitle(int target) {
		String val = null;
		if (target == WorkbookCtrl.HEADER_TYPE_ROW) {
			val = Labels.getLabel("header.rowHeight");
		} else {
			val = Labels.getLabel("header.columnWidth");
		}
		_headerSizeDialog.setTitle(val);
		title.setValue(val + ": ");	
	}

	public void onOK$headerSize() {
		setHeaderSize();
	}
	
	public void onClick$okBtn() {
		setHeaderSize();
	}
	
	private void setHeaderSize() {
		WorkbookCtrl bookCtrl = Zssapp.getDesktopWorkbenchContext(self).getWorkbookCtrl();
		if (headerType ==  WorkbookCtrl.HEADER_TYPE_ROW)
			bookCtrl.setRowHeightInPx(headerSize.getValue(), selection); 
		else
			bookCtrl.setColumnWidthInPx(headerSize.getValue(), selection);
		_headerSizeDialog.fireOnClose(null);
	}
}
