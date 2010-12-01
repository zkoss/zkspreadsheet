/* CellMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 3:06:55 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;
/**
 * @author Sam
 *
 */
public class CellMenupopup extends Menupopup implements IdSpace {

	/* Components */
	private Menuitem cut;
	private Menuitem copy;
	private Menuitem paste;
	private Menuitem pasteSpecial;
	
	//private Menuitem insert;
	private Menuitem shiftCellRight;
	private Menuitem shiftCellDown;
	private Menuitem insertEntireRow;
	private Menuitem insertEntireColumn;

	//private Menuitem delete;
	private Menuitem shiftCellLeft;
	private Menuitem shiftCellUp;
	private Menuitem deleteEntireRow;
	private Menuitem deleteEntireColumn;
	
	private Menuitem clearContent;
	private Menuitem clearStyle;
	
	//TODO: not implement yet
	//private Menuitem filter
	private Menuitem sortAscending;
	private Menuitem sortDescending;
	private Menuitem customSort;
	//TODO: not implement yet
	//private Menuitem comment
	
	//Others
	private Menuitem formula;
	//TODO: not implement yet
	//private Menuitem chart;
	//private Menuitem pic;
	private Menuitem format;
	private Menuitem hyperlink;
	//TODO: not implement yet
	//private Menuitem removeHyperlink;
	
	public CellMenupopup() {
		Executions.createComponents(Consts._CellMenupopup_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}

	public void onOpen() {
		getDesktopWorkbookContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onClick$cut() {
		getDesktopWorkbookContext().getWorkbookCtrl().cutSelection();
	}
	
	public void onClick$copy() {
		getDesktopWorkbookContext().getWorkbookCtrl().copySelection();
	}
	
	public void onClick$paste() {
		getDesktopWorkbookContext().getWorkbookCtrl().pasteSelection();
	}
	
	public void onClick$pasteSpecial() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openPasteSpecialDialog();
	}

	public void onClick$shiftCellRight() {
		getDesktopWorkbookContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_RIGHT);
	}
	
	public void onClick$shiftCellDown() {
		getDesktopWorkbookContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_DOWN);
	}
	
	public void onClick$insertEntireRow() {
		getDesktopWorkbookContext().getWorkbookCtrl().insertRowAbove();	
	}
	
	public void onClick$insertEntireColumn() {
		getDesktopWorkbookContext().getWorkbookCtrl().insertColumnLeft();
	}
	
	public void onClick$shiftCellLeft() {
		getDesktopWorkbookContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_LEFT);
	}
	
	public void onClick$shiftCellUp() {
		getDesktopWorkbookContext().getWorkbookCtrl().shiftCell(DesktopWorkbenchContext.SHIFT_CELL_UP);
	}
	
	public void onClick$deleteEntireRow() {
		getDesktopWorkbookContext().getWorkbookCtrl().deleteRow();
	}

	public void onClick$deleteEntireColumn() {
		getDesktopWorkbookContext().getWorkbookCtrl().deleteColumn();
	}
	
	public void onClick$clearContent() {
		getDesktopWorkbookContext().getWorkbookCtrl().clearSelectionContent();
	}
	
	public void onClick$clearStyle() {
		getDesktopWorkbookContext().getWorkbookCtrl().clearSelectionStyle();
	}

	public void onClick$sortAscending() {
		getDesktopWorkbookContext().getWorkbookCtrl().sort(false);
	}
	
	public void onClick$sortDescending() {
		getDesktopWorkbookContext().getWorkbookCtrl().sort(true);
	}
	
	public void onClick$customSort() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openCustomSortDialog();
	}
	
	public void onClick$formula() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openInsertFormulaDialog();
	}

	public void onClick$format() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openFormatNumberDialog();
	}
	
	public void onClick$hyperlink() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openHyperlinkDialog();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbookContext() {
		return DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop());
	}
}