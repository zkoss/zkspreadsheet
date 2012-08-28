/* RenameSheetCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 6:22:52 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ctrl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Button;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * @author sam
 *
 */
public class RenameSheetCtrl extends GenericForwardComposer {
	
	private final static String KEY_ARG_SHEET_NAME = "org.zkoss.zss.app.ctrl.renameSheetCtrl.sheetName";
	
	private Dialog _renameSheetDialog;
	private Button confirmRenameBtn;
	private Textbox sheetNameTB;
	
	/**
	 * @param originalSheetName
	 * @return
	 */
	public static Map newArg(String originalSheetName) {
		HashMap<String, Object> arg = new HashMap<String, Object>(1);
		arg.put(KEY_ARG_SHEET_NAME, originalSheetName);
		return arg;
	}

	public void onOpen$_renameSheetDialog(ForwardEvent event) {
		Map arg = (Map) event.getOrigin().getData();
		sheetNameTB.setText((String)arg.get(KEY_ARG_SHEET_NAME));
		sheetNameTB.focus();
		_renameSheetDialog.setMode(Window.MODAL);
	}
	
	public void onOK$sheetNameTB() {
		rename();
	}
	
	public void onClick$confirmRenameBtn() {
		rename();
	}
	
	private void rename() {
		String sheetName = sheetNameTB.getText();
		if (sheetName == null || sheetName == "") {
			Messagebox.show("invalid sheet name");
			return;
		}
		DesktopWorkbenchContext bookContent = Zssapp.getDesktopWorkbenchContext(self);
		bookContent.getWorkbookCtrl().renameSelectedSheet(sheetName);
		bookContent.fireRefresh();
		
		_renameSheetDialog.fireOnClose(null);
	}
}