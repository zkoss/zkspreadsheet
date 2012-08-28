/* ImportFileWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 2, 2010 12:53:23 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.file;

import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkspaceContext;
import org.zkoss.zul.Button;
import org.zkoss.zul.Label;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listhead;
import org.zkoss.zul.Listheader;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Radio;
import org.zkoss.zul.Radiogroup;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class ImportFileWindowCtrl extends GenericForwardComposer  {
	
	private Dialog _importFileDialog;
	private Label supportedFormat;
	private Radiogroup importOption;
	private Radio createNew;
	
	//TODO: not implement yet
	private Radio insertSheetsToEnd;
	//TODO: not implement yet
	private Radio replaceCurrent;
	
	//TODO: provide search bar for filter file name
	/*All spreadsheet file name list*/
	private Listbox allFilesListbox;
	
	private Button uploadBtn;
	
	private Menupopup fileMenu;
	private Menuitem openFileMenuitem;
	
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		initSupportFormat();
		initImportOption();
		initFileListbox();
	}
	
	public void onOpen$_importFileDialog() {
		_importFileDialog.setMode(Window.MODAL);
		Map<String, SpreadSheetMetaInfo> metaInfos = SpreadSheetMetaInfo.getMetaInfos();
		if (metaInfos == null || metaInfos.isEmpty())
			return;
		allFilesListbox.setModel(new ListModelList(metaInfos.values()));
	}

	private void initSupportFormat() {	
		String val = "";
		String[] supporedFormats = FileHelper.getSupportedFormat();
		for (int i = 0; i < supporedFormats.length; i++) {
			val += i == 0 ? supporedFormats[i] : ", " + supporedFormats[i];
		}
		//TODO: use I18n
		supportedFormat.setValue("Supported formats: " + val);
	}

	/**
	 * Initialize all spreadsheet file name as a list 
	 */
	private void initFileListbox() {

		//TODO: move this to become a component, re-use in here and fileListOpen.zul
		Listhead listhead = new Listhead();
		Listheader filenameHeader = new Listheader("File");
		filenameHeader.setHflex("2");
		filenameHeader.setParent(listhead);
		
		Listheader dateHeader = new Listheader("Date");
		dateHeader.setHflex("1");
		dateHeader.setParent(listhead);
		allFilesListbox.appendChild(listhead);
		
		allFilesListbox.setItemRenderer(new ListitemRenderer() {

			@Override
			public void render(Listitem item, Object data, int index)
					throws Exception {
				final SpreadSheetMetaInfo info = (SpreadSheetMetaInfo)data;
				item.setValue(info);
				item.appendChild(new Listcell(info.getFileName()));
				item.appendChild(new Listcell(info.getFormatedImportDateString()));
				//TODO: use I18n
				item.setContext(fileMenu);
				item.addEventListener(Events.ON_DOUBLE_CLICK, new EventListener() {
					
					@Override
					public void onEvent(Event evt) throws Exception {
						DesktopWorkbenchContext workbenchCtrl = getDesktopWorkbenchContext();
						workbenchCtrl.getWorkbookCtrl().openBook(info);
						workbenchCtrl.fireWorkbookChanged();
						_importFileDialog.fireOnClose(null);
					}
				});
			}
		});
	}

	/**
	 * Initialize import option, sets create new file as default
	 */
	private void initImportOption() {
		importOption.setSelectedItem(createNew);
		//TODO: implement other option
	}
	
	public void onClick$openFileMenuitem() {
		getDesktopWorkbenchContext().getWorkbookCtrl().openBook((SpreadSheetMetaInfo)allFilesListbox.getSelectedItem().getValue());
		getDesktopWorkbenchContext().fireWorkbookChanged();
		_importFileDialog.fireOnClose(null);
	}
	
	public void onFileUpload(ForwardEvent event) {
		try {
			WorkspaceContext.getInstance(desktop).store(((UploadEvent) event.getOrigin()).getMedia());
		} catch (UnsupportedSpreadSheetFileException e) {
			//TODO: use I18n to replace message string
			Messagebox.show("Please import xls or xlsx file");
		}
		allFilesListbox.setModel(new ListModelList(SpreadSheetMetaInfo.getMetaInfos().values()));	
	}
	
	private DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
}