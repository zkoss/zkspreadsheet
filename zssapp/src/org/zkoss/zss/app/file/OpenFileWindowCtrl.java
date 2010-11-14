/* OpenFileWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 3, 2010 7:26:02 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.file;

import java.util.Map;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ZssappComponents;
import org.zkoss.zss.ui.Spreadsheet;
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
import org.zkoss.zul.Toolbarbutton;

/**
 * @author Sam
 *
 */
public class OpenFileWindowCtrl extends GenericForwardComposer {
	
	//TODO: rename this
	private Listbox filesListbox;
	
	private Spreadsheet ss;
	
	private Menupopup fileMenu;
	private Menuitem openFileMenuitem;
	
	private Toolbarbutton uploadAndOpenBtn;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		ss = ZssappComponents.getSpreadsheetFromArg();
		initFileListbox();
	}

	private void initFileListbox() {
		Map<String, SpreadSheetMetaInfo> metaInfos = SpreadSheetMetaInfo.getMetaInfos();
		if (metaInfos == null || metaInfos.isEmpty()) {
			System.out.println("metaInfos is empty");
			return;
		}
		
		//TODO: move this to become a component, re-use in here and fileListOpen.zul
		Listhead listhead = new Listhead();
		Listheader filenameHeader = new Listheader("File");
		filenameHeader.setHflex("2");
		filenameHeader.setParent(listhead);
		
		Listheader dateHeader = new Listheader("Date");
		dateHeader.setHflex("1");
		dateHeader.setParent(listhead);
		filesListbox.appendChild(listhead);
		
		filesListbox.setModel(new ListModelList(metaInfos.values()));
		filesListbox.setItemRenderer(new ListitemRenderer() {
			
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
						Zssapp.redrawSheets(ss);
						((Component)spaceOwner).detach();
					}
				});
			}
		});
	}
	
	public void onUpload$uploadAndOpenBtn(UploadEvent event) {
		//TODO: throw IO exception
		try {
			Media media = event.getMedia();
			FileHelper.store(media);

			SpreadSheetMetaInfo info = (SpreadSheetMetaInfo)SpreadSheetMetaInfo.getMetaInfos().get(media.getName());
			if (info != null)
				FileHelper.openSpreadsheet(ss, info);
			Zssapp.redrawSheets(ss);
			//else
			//TODO: throw Io exception message
			((Component)spaceOwner).detach();
		} catch (UnsupportedSpreadSheetFileException e) {
			try {
				//TODO: I18n
				Messagebox.show("Import file format " + event.getMedia().getName() + " not supported");
			} catch (InterruptedException e1) {
			}
		}
	}

	public void onClick$openFileMenuitem() {
		FileHelper.openSpreadsheet(ss, 
				(SpreadSheetMetaInfo)filesListbox.getSelectedItem().getValue());
		Zssapp.redrawSheets(ss);
		((Component)spaceOwner).detach();
	}
}
