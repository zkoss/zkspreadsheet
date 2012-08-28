/* OpenFileWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 3, 2010 7:26:02 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.file;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkspaceContext;
import org.zkoss.zul.Button;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listhead;
import org.zkoss.zul.Listheader;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class OpenFileWindowCtrl extends GenericForwardComposer {
	
	private Dialog _openFileDialog;
	private Listbox filesListbox;
	
	private Button uploadBtn;

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		initFileListbox();
		uploadBtn.setDisabled(!FileHelper.hasImportPermission());
	}
	
	public void onOpen$_openFileDialog() {
		filesListbox.setModel(new ListModelList(
				WorkspaceContext.getInstance(desktop).getMetainfos()));
		_openFileDialog.setMode(Window.MODAL);
	}

	private void initFileListbox() {
		//TODO: move this to become a component, re-use in here and fileListOpen.zul
		Listhead listhead = new Listhead();
		Listheader filenameHeader = new Listheader("File");
		filenameHeader.setHflex("2");
		filenameHeader.setParent(listhead);
		
		Listheader dateHeader = new Listheader("Date");
		dateHeader.setHflex("1");
		dateHeader.setParent(listhead);
		filesListbox.appendChild(listhead);
		
		filesListbox.setItemRenderer(new ListitemRenderer() {
			
			public void render(Listitem item, Object obj) throws Exception {
				final SpreadSheetMetaInfo info = (SpreadSheetMetaInfo)obj;
				item.setValue(info);
				item.appendChild(new Listcell(info.getFileName()));
				item.appendChild(new Listcell(info.getFormatedImportDateString()));
				//TODO: use I18n
				item.addEventListener(Events.ON_DOUBLE_CLICK, new EventListener() {
					public void onEvent(Event evt) throws Exception {
						getDesktopWorkbenchContext().getWorkbookCtrl().openBook(info);
						getDesktopWorkbenchContext().fireWorkbookChanged();
						_openFileDialog.fireOnClose(null);
					}
				});
			}

			@Override
			public void render(Listitem item, Object data, int index)
					throws Exception {
				render(item, data);
			}
		});
	}
	
	public void onUpload$uploadBtn(UploadEvent event) {
		//TODO: throw IO exception
		try {
			getDesktopWorkbenchContext().getWorkbookCtrl().
				openBook(WorkspaceContext.getInstance(desktop).store(event.getMedia()));
			getDesktopWorkbenchContext().fireWorkbookChanged();
			_openFileDialog.fireOnClose(null);
		} catch (UnsupportedSpreadSheetFileException e) {
			Messagebox.show("Import file format " + event.getMedia().getName() + " not supported");
		}
	}
	
	private DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
}