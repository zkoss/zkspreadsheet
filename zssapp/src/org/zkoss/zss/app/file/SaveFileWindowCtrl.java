/* SaveFileWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 17, 2010 5:23:32 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

 */
package org.zkoss.zss.app.file;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Comboitem;
import org.zkoss.zul.ComboitemRenderer;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * @author Sam
 * 
 */
public class SaveFileWindowCtrl extends GenericForwardComposer {

	Dialog _saveFileDialog;
	Combobox fileFormat;
	Textbox fileName;
	Button okBtn;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);

		fileFormat.setReadonly(true);
		fileFormat.setItemRenderer(new ComboitemRenderer() {
			public void render(Comboitem item, Object data) throws Exception {
				item.setLabel(data.toString());
			}

			@Override
			public void render(Comboitem item, Object data, int index)
					throws Exception {
				render(item, data);
			}
		});
		fileFormat.addEventListener("onAfterRender", new EventListener() {
			public void onEvent(Event event) throws Exception {
				if (fileFormat.getModel().getSize() > 0)
					fileFormat.setSelectedIndex(0);
			}
		});
	}
	
	public void onOpen$_saveFileDialog() {
		fileFormat.setModel(new ListModelList(FileHelper.getSupportedFormat()));

		String src = getDesktopWorkbenchContext().getWorkbookCtrl().getSrc();
		if ("Untitled".lastIndexOf(src) >= 0) {
			fileName.setValue("Book1");
		}
		_saveFileDialog.setMode(Window.MODAL);
	}

	public void onOK$fileName() {
		save();
	}

	public void onClick$okBtn() {
		save();
	}

	private void save() {
		if (!"".equals(fileName.getText())) {
			getDesktopWorkbenchContext().getWorkbookCtrl().setSrcName(
					fileName.getText() + "." + fileFormat.getSelectedItem().getLabel());
			getDesktopWorkbenchContext().getWorkbookCtrl().save();
			getDesktopWorkbenchContext().fireWorkbookSaved();
			_saveFileDialog.fireOnClose(null);
		} else
			Messagebox.show("File name can not be empty");
	}

	private DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
}
