package org.zkoss.zss.app.ctrl;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zul.Button;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;

/**
 * TODO: use inline-edit if possible
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
		MainWindowCtrl.getInstance().renameSheet(sheetName);
		((Component)spaceOwner).detach();
	}
}
