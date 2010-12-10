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

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class EditMenu extends Menu implements IdSpace {
	
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
		Executions.createComponents(Consts._EditMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');

		getDesktopWorkbenchContext().addEventListener(Consts.ON_SHEET_OPEN, new EventListener() {
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
		getDesktopWorkbenchContext().getWorkbookCtrl().cutSelection();
	}
	
	public void onClick$copy() {
		getDesktopWorkbenchContext().getWorkbookCtrl().copySelection();
	}
	
	public void onClick$paste() {
		getDesktopWorkbenchContext().getWorkbookCtrl().pasteSelection();
	}
	
	public void onClick$clearContent() {
		getDesktopWorkbenchContext().getWorkbookCtrl().clearSelectionContent();
	}
	
	public void onClick$clearStyle() {
		getDesktopWorkbenchContext().getWorkbookCtrl().clearSelectionStyle();
	}
	
	public void onClick$clearAll() {
		getDesktopWorkbenchContext().getWorkbookCtrl().clearSelectionContent();
		getDesktopWorkbenchContext().getWorkbookCtrl().clearSelectionStyle();
	}
	
	public void onClick$shiftCellRight() {
		getDesktopWorkbenchContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_RIGHT);
	}
	
	public void onClick$shiftCellDown() {
		getDesktopWorkbenchContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_DOWN);
	}
	
	public void onClick$insertEntireRow() {
		getDesktopWorkbenchContext().getWorkbookCtrl().insertRowAbove();
	}
	
	public void onClick$insertEntireColumn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().insertColumnLeft();
	}
	
	public void onClick$shiftCellLeft() {
		getDesktopWorkbenchContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_LEFT);
	}
	
	public void onClick$shiftCellUp() {
		getDesktopWorkbenchContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_UP);
	}
	
	public void setDisabled(boolean disabled) {
		//TODO: undo, redo is not implement yet
		undo.setDisabled(true);
		redo.setDisabled(true);
		
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
		getDesktopWorkbenchContext().getWorkbookCtrl().deleteRow();
	}
	
	public void onClick$deleteEntireColumn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().deleteColumn();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop());
	}
	
	public void onOpen$editMenupopup() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
}