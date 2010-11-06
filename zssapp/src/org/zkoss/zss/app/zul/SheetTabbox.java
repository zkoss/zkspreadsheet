package org.zkoss.zss.app.zul;

import java.util.HashMap;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.util.Composer;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.app.ctrl.RenameSheetCtrl;
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
 * SheetTabbox is responsible for switch sheets
 * @author sam
 *
 */
public class SheetTabbox extends Div implements Composer {
	
	private final static String URI = "/macros/sheetTabbox.zul";
	
	private Tabbox tabbox;
	
	private Tabs tabs;
	
	private Menupopup sheetContextMenu;
	
	private Menuitem shiftSheetLeft;
	
	private Menuitem shiftSheetRight;
	
	private Menuitem deleteSheet;
	
	private Menuitem renameSheet;
	
	private Menuitem copySheet;
	
	private Spreadsheet ss;
	
	public SheetTabbox() {
		Executions.createComponents(URI, this, null);
	}

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		Components.wireVariables(comp, (Object)this);
		Components.addForwards(comp, (Object)this, '$');

		ss = MainWindowCtrl.getInstance().getSpreadsheet();
		redraw();
	}
	
	public void onSelect$tabbox() {
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
		//TODO: show message if fail to shift
		MainWindowCtrl.getInstance().shiftSheetLeft();
	}
	
	public void onClick$shiftSheetRight() {
		//TODO: show message if fail to shift
		MainWindowCtrl.getInstance().shiftSheetRight();
	}
	
	public void onClick$deleteSheet() {
		//TODO: show message if fail to shift
		MainWindowCtrl.getInstance().deleteSheet();
	}
	
	public void onClick$renameSheet() {
		HashMap<String, String> arg = new HashMap<String, String>();
		arg.put(RenameSheetCtrl.KEY_ARG_SHEET_NAME, getCurrenSheet());
		Executions.createComponents("/menus/tab/rename.zul", null, arg);
	}
	
	public void onClick$copySheet() {
		throw new UiException("cop sheet not implement yet");
	}
}
