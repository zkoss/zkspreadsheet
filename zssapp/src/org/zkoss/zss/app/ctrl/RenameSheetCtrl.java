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

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zul.Button;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;

/**
 * @author sam
 *
 */
public class RenameSheetCtrl extends GenericForwardComposer {
	
	public final static String KEY_ARG_SHEET_NAME = "org.zkoss.zss.app.ctrl.renameSheetCtrl.sheetName";
	
	private Button confirmRenameBtn;
	
	private Textbox sheetNameTB;

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);

		sheetNameTB.setText((String)Executions.getCurrent().getArg().get(KEY_ARG_SHEET_NAME));
	}
	
	public void onClick$confirmRenameBtn() {
		String sheetName = sheetNameTB.getText();
		if (sheetName == null || sheetName == "") {
			try {
				Messagebox.show("invalid sheet name");
			} catch (InterruptedException e) {
			}
			return;
		}
		DesktopWorkbenchContext bookContent = DesktopWorkbenchContext.getInstance(desktop);
		bookContent.getWorkbookCtrl().renameSelectedSheet(sheetName);
		bookContent.fireRefresh();
		
		self.detach();
	}
}