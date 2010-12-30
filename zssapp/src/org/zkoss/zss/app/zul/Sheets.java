/* Sheets.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Apr 12, 2010 3:47:56 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.SheetVisitor;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Div;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Tab;
import org.zkoss.zul.Tabbox;
import org.zkoss.zul.Tabs;

/**
 * Sheets is responsible for switch sheets
 * @author sam
 *
 */
public class Sheets extends Div implements ZssappComponent, IdSpace {

	private Tabbox tabbox;
	private Tabs tabs;
	
	private Menupopup sheetContextMenu;
	private Menuitem shiftSheetLeft;
	private Menuitem shiftSheetRight;
	private Menuitem deleteSheet;
	private Menuitem renameSheet;
	private Menuitem copySheet;
	
	private Spreadsheet ss;
	
	public Sheets() {
		Executions.createComponents(Consts._SheetPanel_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void onSelect$tabbox() {
		getDesktopWorkbenchContext().getWorkbookCtrl().setSelectedSheet(tabbox.getSelectedTab().getLabel());
		getDesktopWorkbenchContext().fireSheetChanged();
	}
	
	/**
	 * Sets the current sheet name of the tabbox
	 * @param index
	 */
	public void setCurrentSheet(int index) {
		tabbox.setSelectedIndex(index);
		getDesktopWorkbenchContext().getWorkbookCtrl().setSelectedSheet(tabbox.getSelectedTab().getLabel());
	}
	
	/**
	 * Returns the current sheet name of tabbox
	 * @return
	 */
	public String getCurrenSheet() {
		return tabbox.getSelectedTab().getLabel();
	}
	
	/**
	 * Add a sheet name to tabbox
	 * @param name
	 */
	public void addSheet(String name) {
		Tab tab = new Tab(name);
		tab.setContext(sheetContextMenu);
		tab.setParent(tabs);
	}
	
	/**
	 * Redraw sheet names of spreadsheet
	 */
	public void redraw() {
		 final int seldIndex = tabbox.getSelectedTab() != null ?
				 tabbox.getSelectedTab().getIndex() : -1;
		
		tabs.getChildren().clear();

		Utils.visitSheets(ss.getBook(), new SheetVisitor(){
			@Override
			public void handle(Sheet sheet) {
				final Tab tab = new Tab(sheet.getSheetName());
				tab.addEventListener(org.zkoss.zk.ui.event.Events.ON_RIGHT_CLICK, new EventListener() {
					public void onEvent(Event event) throws Exception {						
						
						if (tabbox.getSelectedTab().getLabel() != tab.getLabel())
							setSelectedTab(tab);

						MouseEvent evt = (MouseEvent)event;
						sheetContextMenu.open(evt.getPageX(), evt.getPageY());
					}
				});
				tab.setParent(tabs);
				if (seldIndex == tab.getIndex())
					setSelectedTab(tab);
			}});
	}

	private void setSelectedTab(Tab tab) {
		tabbox.setSelectedTab(tab);
		getDesktopWorkbenchContext().getWorkbookCtrl().setSelectedSheet(tab.getLabel());
		getDesktopWorkbenchContext().fireSheetChanged();
	}
	
	/**
	 * Clear all tabs
	 */
	public void clear() {
		tabs.getChildren().clear();
	}
	
	public void onClick$shiftSheetLeft() {
		int newIdx = SheetHelper.shiftSheetLeft(checkNotNull(ss, "Spreadsheet is null"));
		if (newIdx >= 0) {
			tabbox.setSelectedIndex(newIdx);
			redraw();
			setCurrentSheet(newIdx);
		} else {
			//TODO: show error message
		}
	}
	
	public void onClick$shiftSheetRight() {
		int newIdx = SheetHelper.shiftSheetRight(checkNotNull(ss, "Spreadsheet is null"));
		if (newIdx >= 0) {
			tabbox.setSelectedIndex(newIdx);
			redraw();
			setCurrentSheet(newIdx);
		} else {
			//TODO: show error message
		}
	}
	
	public void onClick$deleteSheet() {
		//TODO: show message if fail to shift
		int newIdx = SheetHelper.deleteSheet(checkNotNull(ss, "Spreadsheet is null"));
		if (newIdx >= 0) {
			tabbox.setSelectedIndex(newIdx);
			redraw();
			setCurrentSheet(newIdx);
		} else {
			try {
				Messagebox.show("A workbook must contain at least one visible worksheet");
			} catch (InterruptedException e) {
			}
		}
	}
	
	public void onClick$renameSheet() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openRenameSheetDialog(getCurrenSheet());
	}
	
	public void onClick$copySheet() {
		throw new UiException("cop sheet not implement yet");
	}

	@Override
	public void unbindSpreadsheet() {
		//TODO: unbind event
	}

	@Override
	public void bindSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
}