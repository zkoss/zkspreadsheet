/* EditMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 5:28:23 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.app.zul.ctrl.DesktopSheetContext;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class EditMenu extends Menu implements ZssappComponent, IdSpace {

	private final static String URI = "~./zssapp/html/menu/editMenu.zul";
	
	private Menupopup editMenupopup;
	
	/*TODO: not implement yet*/
	private Menuitem undo;
	private Menuitem redo;
	
	private Menuitem cut;
	private Menuitem copy;
	private Menuitem paste;
	
	private Menuitem clearContent;
	private Menuitem clearStyle;
	private Menuitem clearAll;
	
	private Menuitem shiftCellRight;
	private Menuitem shiftCellDown;
	private Menuitem insertEntireRow;
	private Menuitem insertEntireColumn;

	private Menuitem shiftCellLeft;
	private Menuitem shiftCellUp;
	private Menuitem deleteEntireRow;
	private Menuitem deleteEntireColumn;
	
	private Spreadsheet ss;
	
	public EditMenu() {
		Executions.createComponents(URI, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
		
		DesktopSheetContext.getInstance(this.getDesktop()).
		addEventListener(Consts.ON_SHEET_OPEN, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				setDisabled(!(Boolean)event.getData());
			}
		});
	}
	
	public void onClick$undo() {
		throw new UiException("undo not implement yet");
	}
	
	public void onClick$redo() {
		throw new UiException("redo not implement yet");
	}
	
	public void onClick$cut() {
		EditHelper.doCut(ss);
	}
	
	public void onClick$copy() {
		EditHelper.doCopy(ss);
	}
	
	public void onClick$paste() {
		EditHelper.doPaste(ss);
	}
	
	public void onClick$clearContent() {
		CellHelper.clearContent(ss, SheetHelper.getSpreadsheetMaxSelection(ss));
	}
	
	public void onClick$clearStyle() {
		CellHelper.clearStyle(ss, SheetHelper.getSpreadsheetMaxSelection(ss));
	}
	
	public void onClick$clearAll() {
		CellHelper.clearContent(ss, SheetHelper.getSpreadsheetMaxSelection(ss));
		CellHelper.clearStyle(ss, SheetHelper.getSpreadsheetMaxSelection(ss));
	}
	
	public void onClick$shiftCellRight() {
		CellHelper.shiftCellRight(ss.getSelectedSheet(), 
				ss.getSelection().getTop(), 
				ss.getSelection().getLeft());
	}
	
	public void onClick$shiftCellDown() {
		CellHelper.shiftCellDown(ss.getSelectedSheet(), 
				ss.getSelection().getTop(), 
				ss.getSelection().getLeft());
	}
	
	public void onClick$insertEntireRow() {
		CellHelper.shiftEntireRowDown(ss.getSelectedSheet(), 
				ss.getSelection().getTop(), 
				ss.getSelection().getLeft());
	}
	
	public void onClick$insertEntireColumn() {
		CellHelper.shiftEntireColumnRight(ss.getSelectedSheet(), 
				ss.getSelection().getTop(), 
				ss.getSelection().getLeft());
	}
	
	public void onClick$shiftCellLeft() {
		CellHelper.shiftCellLeft(ss.getSelectedSheet(), 
				ss.getSelection().getTop(), 
				ss.getSelection().getLeft());
	}
	
	public void onClick$shiftCellUp() {
		CellHelper.shiftCellUp(ss.getSelectedSheet(), 
				ss.getSelection().getTop(), 
				ss.getSelection().getLeft());
	}
	
	public void setDisabled(boolean disabled) {
		undo.setDisabled(disabled);
		redo.setDisabled(disabled);
		
		cut.setDisabled(disabled);
		copy.setDisabled(disabled);
		paste.setDisabled(disabled);
		
		clearContent.setDisabled(disabled);
		clearStyle.setDisabled(disabled);
		clearAll.setDisabled(disabled);
		
		shiftCellRight.setDisabled(disabled);
		shiftCellDown.setDisabled(disabled);
		insertEntireRow.setDisabled(disabled);
		insertEntireColumn.setDisabled(disabled);
		
		shiftCellLeft.setDisabled(disabled);
		shiftCellUp.setDisabled(disabled);
		deleteEntireRow.setDisabled(disabled);
		deleteEntireColumn.setDisabled(disabled);
	}
	
	public void onClick$deleteEntireRow() {
		Sheet sheet = ss.getSelectedSheet();
		Rect rect = ss.getSelection();
		int tRow = rect.getTop();
		int lCol = rect.getLeft();
		CellHelper.shiftEntireRowUp(ss.getSelectedSheet(), tRow, lCol);
	}
	
	public void onClick$deleteEntireColumn() {
		Rect rect = ss.getSelection();
		int tRow = rect.getTop();
		int lCol = rect.getLeft();
		CellHelper.shiftEntireColumnLeft(ss.getSelectedSheet(), tRow, lCol);
	}
	
	@Override
	public void bindSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
		editMenupopup.setWidgetListener(Events.ON_OPEN, "this.$f('" + ss.getId() + "', true).focus(false);");
	}

	@Override
	public void unbindSpreadsheet() {
		// TODO Auto-generated method stub
		
	}

}