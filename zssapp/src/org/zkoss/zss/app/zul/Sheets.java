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

import java.util.List;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
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
public class Sheets extends Div implements IdSpace {

	private Tabbox tabbox;
	private Tabs tabs;
	
	private Menupopup sheetContextMenu;
	private Menuitem shiftSheetLeft;
	private Menuitem shiftSheetRight;
	private Menuitem deleteSheet;
	private Menuitem renameSheet;
	private Menuitem copySheet;
	private Menuitem protectSheet;
	
	public Sheets() {
		Executions.createComponents(Consts._SheetPanel_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void onCreate() {
		final DesktopWorkbenchContext workbenchContext= getDesktopWorkbenchContext();
		workbenchContext.addEventListener(Consts.ON_SHEET_CHANGED, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				protectSheet.setChecked(workbenchContext.getWorkbookCtrl().isSheetProtect());
			}
		});		
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

		List<String> names = getDesktopWorkbenchContext().getWorkbookCtrl().getSheetNames();
		for (String name : names) {
			final Tab tab = new Tab(name);
			tab.addEventListener(org.zkoss.zk.ui.event.Events.ON_RIGHT_CLICK, new EventListener() {
			public void onEvent(Event event) throws Exception {						
				if (tabbox.getSelectedTab().getLabel() != tab.getLabel())
					setSelectedTab(tab);
	
				protectSheet.setChecked(getDesktopWorkbenchContext().getWorkbookCtrl().isSheetProtect());
				MouseEvent evt = (MouseEvent)event;
				sheetContextMenu.open(evt.getPageX(), evt.getPageY());
			}});
			tab.setParent(tabs);
			if (seldIndex == tab.getIndex()) {
				setSelectedTabDirectly(tab);
			}
		}
	}

	private void setSelectedTab(Tab tab) {
		tabbox.setSelectedTab(tab);
		getDesktopWorkbenchContext().getWorkbookCtrl().setSelectedSheet(tab.getLabel());
		getDesktopWorkbenchContext().fireSheetChanged();
	}
	
	private void setSelectedTabDirectly(Tab tab) {
		tabbox.setSelectedTab(tab);
	}
	
	/**
	 * Clear all tabs
	 */
	public void clear() {
		tabs.getChildren().clear();
	}
	
	public void onClick$shiftSheetLeft() {
		int newIdx = getDesktopWorkbenchContext().getWorkbookCtrl().shiftSheetLeft();
		if (newIdx >= 0) {
			tabbox.setSelectedIndex(newIdx);
			redraw();
			setCurrentSheet(newIdx);
		} else {
			//TODO: show error message
		}
	}
	
	public void onClick$shiftSheetRight() {
		int newIdx = getDesktopWorkbenchContext().getWorkbookCtrl().shiftSheetRight();
		if (newIdx >= 0) {
			tabbox.setSelectedIndex(newIdx);
			redraw();
			setCurrentSheet(newIdx);
		} else {
			//TODO: show error message
		}
	}
	
	public void onClick$deleteSheet() {
		int newIdx = getDesktopWorkbenchContext().getWorkbookCtrl().deleteSheet();
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
	
	public void onCheck$protectSheet() {
		getDesktopWorkbenchContext().getWorkbookCtrl().protectSheet(
			protectSheet.isChecked() ? "" : null);
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
}