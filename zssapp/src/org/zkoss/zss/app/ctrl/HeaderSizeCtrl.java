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

import java.util.Map;

import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zul.Button;
import org.zkoss.zul.Intbox;
import org.zkoss.zul.Label;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class HeaderSizeCtrl extends GenericForwardComposer {
	
	/* header type to set, specify column or row */
	public final static String KEY_HEADER_TYPE = "org.zkoss.zss.app.ctrl.HeaderSizeCtrl.headerType";
	private Integer headerType;
	
	/* header original value */
	public final static String KEY_HEADER_SIZE = "org.zkoss.zss.app.ctrl.HeaderSizeCtrl.headerSize";
	private Integer size;
	
	/* dialog window */
	private Window headerSizeWin;
	
	/* header title specify Column width or Row height */
	private Label title;
	
	/* header size of the target */
	private Intbox headerSize;
	
	private Button okBtn;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		Map arg = Executions.getCurrent().getArg();
		headerType = (Integer)arg.get(KEY_HEADER_TYPE);
		initTitle(headerType);
		size = (Integer)arg.get(KEY_HEADER_SIZE);
		headerSize.setValue(size);
	}
	
	private void initTitle(int target) {
		String val = null;
		if (target == WorkbookCtrl.HEADER_TYPE_ROW) {
			val = Labels.getLabel("header.rowHeight");
		} else {
			val = Labels.getLabel("header.columnWidth");
		}
		headerSizeWin.setTitle(val);
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
			bookCtrl.setRowHeightInPx(headerSize.getValue()); 
		else
			bookCtrl.setColumnWidthInPx(headerSize.getValue());
		self.detach();
	}
}
