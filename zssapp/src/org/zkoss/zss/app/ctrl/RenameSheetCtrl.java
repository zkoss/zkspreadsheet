package org.zkoss.zss.app.ctrl;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.ui.Spreadsheet;
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
	
	private Spreadsheet ss;

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);

		ss = (Spreadsheet)Executions.getCurrent().getArg().get("spreadsheet");
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
		int idx = SheetHelper.renameSheet(ss, sheetName);
		if (idx >= 0) {
			Zssapp.redrawSheets(ss);
			((Component)spaceOwner).detach();
		} else {
			//TODO: show error message
		}
	}
}
