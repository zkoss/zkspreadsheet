package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import java.util.HashMap;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.app.ctrl.RenameSheetCtrl;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.SheetVisitor;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Div;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;
import org.zkoss.zul.Tab;
import org.zkoss.zul.Tabbox;
import org.zkoss.zul.Tabs;

/**
 * SheetsName is responsible for switch sheets
 * @author sam
 *
 */
public class Sheets extends Div implements ZssappComponent, IdSpace {
	
	private final static String URI = "~./zssapp/html/sheets.zul";
	
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
		Executions.createComponents(URI, this, null);
		
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void onSelect$tabbox() {
		//TODO: use MainWindowCtrl as 
		MainWindowCtrl.getInstance().setSelectedSheet(tabbox.getSelectedTab().getLabel());
	}
	
	/**
	 * Sets the current sheet name of the tabbox
	 * @param index
	 */
	public void setCurrentSheet(int index) {
		tabbox.setSelectedIndex(index);
		MainWindowCtrl.getInstance().setSelectedSheet(tabbox.getSelectedTab().getLabel());
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
		tabs.getChildren().clear();
		Utils.visitSheets(ss.getBook(), new SheetVisitor(){

			@Override
			public void handle(Sheet sheet) {
				Tab tab = new Tab(sheet.getSheetName());
				tab.setContext(sheetContextMenu);
				tab.setParent(tabs);
			}});
		//TODO: need invaldate ?
		tabbox.invalidate();
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
			redraw();
			setCurrentSheet(newIdx);
		} else {
			//TODO: show error message
		}
	}
	
	public void onClick$shiftSheetRight() {
		int newIdx = SheetHelper.shiftSheetRight(checkNotNull(ss, "Spreadsheet is null"));
		if (newIdx >= 0) {
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
			redraw();
			setCurrentSheet(newIdx);
		} else {
			//TODO: add a new sheet ?
		}
	}
	
	public void onClick$renameSheet() {
		HashMap arg = ZssappComponents.newSpreadsheetArg(ss);
		arg.put(RenameSheetCtrl.KEY_ARG_SHEET_NAME, getCurrenSheet());
		//TODO: replace with simple inline editing, don't need to use window component
		Executions.createComponents("~./zssapp/html/renameDlg.zul", null, arg);
	}
	
	public void onClick$copySheet() {
		throw new UiException("cop sheet not implement yet");
	}


	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
	}
}