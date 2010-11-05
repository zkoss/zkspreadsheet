package org.zkoss.zss.app;

import java.util.List;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zss.model.Book;
//import org.zkoss.zss.model.Sheet;
//import org.zkoss.zss.model.impl.BookImpl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Popup;
import org.zkoss.zul.Tab;
import org.zkoss.zul.Tabbox;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

public class TabHelper {
	Spreadsheet spreadsheet;

	public TabHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}

	public void dispatcher(String type) {
		if (type.equals("shiftLeft")) {
			onTabShiftLeft();
		} else if (type.equals("shiftRight")) {
			onTabShiftRight();
		} else if (type.equals("delete")) {
			onTabDelete();
		} else if (type.equals("rename")) {
			onTabRename();
		} else if (type.equals("renameOK")) {
			onTabRenameOK();
		}

	}

	public void onTabShiftLeft() {
		Book book = spreadsheet.getBook();
		String name = spreadsheet.getSelectedSheet().getSheetName();
		Sheet sheet = spreadsheet.getSelectedSheet();
		int index = book.getSheetIndex(sheet);
		if (index > 0) {
			book.setSheetOrder(name, index - 1);
			redrawTab(index-1);
		}
		
	}

	public void onTabShiftRight() {
		Book book = spreadsheet.getBook();
		String name = spreadsheet.getSelectedSheet().getSheetName();
		Sheet sheet = spreadsheet.getSelectedSheet();
		int index = book.getSheetIndex(sheet);
		if (index < book.getNumberOfSheets() - 1) {
			book.setSheetOrder(name, index + 1);
			redrawTab(index+1);
		}
		
	}

	public void onTabDelete() {
		Book book = spreadsheet.getBook();
		int index = book.getSheetIndex(spreadsheet.getSelectedSheet());
		book.removeSheetAt(index);
		int sheetCount = book.getNumberOfSheets();
		Tabbox sheetTB = (Tabbox) Path.getComponent("//p1/mainWin/sheetTB");

		redrawTab();

		if (sheetCount > 0) {
			sheetTB.setSelectedIndex(0);
			spreadsheet.setSelectedSheet(sheetTB.getSelectedTab().getLabel());
		} else {
			spreadsheet.invalidate();
			// TODO: after remove all sheets, why is there still sheet1's
			// content?
		}
	}

	public void onTabRename() {
		Window win = (Window) Executions.createComponents(
				"/menus/tab/rename.zul", null, null);
		try {
			win.doModal();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onTabRenameOK() {
		Textbox trw_sheetName = (Textbox) Path
				.getComponent("//p1/tabRenameWin/trw_sheetName");

		String newName = trw_sheetName.getValue();
		Book book = spreadsheet.getBook();
		final Sheet selsheet  = spreadsheet.getSelectedSheet();
		Sheet sheet = book.getSheet(newName);
		if(sheet!=null){
			try {
				Messagebox.show("cannot set sheet name the same as other sheets");
				return;
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}
		final int index = book.getSheetIndex(selsheet);
		book.setSheetName(index, newName);
		
		redrawTab();

		Window tabRenameWin = (Window) Path.getComponent("//p1/tabRenameWin");
		tabRenameWin.detach();
	}
	
	public void redrawTab(){
		redrawTab(-1);
	}
	public void redrawTab(int index) {
		Sheet sheet;
		Tabbox sheetTB = (Tabbox) Path.getComponent("//p1/mainWin/sheetTB");
		sheetTB.getFellow("sheetTabs").getChildren().clear();
		Book book = spreadsheet.getBook();
		int sheetCount = book.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			sheet = (Sheet) book.getSheetAt(i);
			Tab tab = new Tab(sheet.getSheetName());
			Popup popup = (Popup) Path.getComponent("//p1/mainWin/tabPopup");
			tab.setContext(popup);
			tab.setParent(sheetTB.getFellow("sheetTabs"));
		}
		if(index>=0)
			sheetTB.setSelectedIndex(index);
	}
}