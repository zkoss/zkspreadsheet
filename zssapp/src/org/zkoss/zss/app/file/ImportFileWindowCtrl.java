/* ImportFileWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 2, 2010 12:53:23 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.file;

import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.ui.Spreadsheet;
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

/**
 * @author Sam
 *
 */
public class ImportFileWindowCtrl extends GenericForwardComposer  {
	
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
	
	private Spreadsheet ss;
	
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		ss = MainWindowCtrl.getInstance().getSpreadsheet();
		initSupportFormat();
		initImportOption();
		initFileListbox();
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
		Map<String, SpreadSheetMetaInfo> metaInfos = SpreadSheetMetaInfo.getMetaInfos();
		if (metaInfos == null || metaInfos.isEmpty())
			return;
		
		//TODO: move this to become a component, re-use in here and fileListOpen.zul
		Listhead listhead = new Listhead();
		Listheader filenameHeader = new Listheader("File");
		filenameHeader.setHflex("2");
		filenameHeader.setParent(listhead);
		
		Listheader dateHeader = new Listheader("Date");
		dateHeader.setHflex("1");
		dateHeader.setParent(listhead);
		allFilesListbox.appendChild(listhead);
		
		allFilesListbox.setModel(new ListModelList(metaInfos.values()));
		allFilesListbox.setItemRenderer(new ListitemRenderer() {
			
			@Override
			public void render(Listitem item, Object obj) throws Exception {
				final SpreadSheetMetaInfo info = (SpreadSheetMetaInfo)obj;
				item.setValue(info);
				item.appendChild(new Listcell(info.getFileName()));
				item.appendChild(new Listcell(info.getFormatedImportDateString()));
				//TODO: use I18n
				item.setContext(fileMenu);
				item.addEventListener(Events.ON_DOUBLE_CLICK, new EventListener() {
					
					@Override
					public void onEvent(Event evt) throws Exception {
						FileHelper.openSpreadsheet(ss, info);
						MainWindowCtrl.getInstance().initSheetNameTab();
						((Component)spaceOwner).detach();
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
		FileHelper.openSpreadsheet(ss, 
				(SpreadSheetMetaInfo)allFilesListbox.getSelectedItem().getValue());
		MainWindowCtrl.getInstance().initSheetNameTab();
		((Component)spaceOwner).detach();
	}
	
	public void onFileUpload(ForwardEvent event) {
		try {
			FileHelper.store(((UploadEvent) event.getOrigin()).getMedia());
			
		} catch (UnsupportedSpreadSheetFileException e) {
			try {
				//TODO: use I18n to replace message string
				Messagebox.show("Please import xls or xlsx file");
			} catch (InterruptedException ex) {
			}
			
		}
		allFilesListbox.setModel(new ListModelList(SpreadSheetMetaInfo.getMetaInfos().values()));	
	}
}