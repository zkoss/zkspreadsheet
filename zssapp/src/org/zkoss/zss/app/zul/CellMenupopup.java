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
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopSheetContext;
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
	
	
//	private Spreadsheet ss;
	
	public CellMenupopup() {
		Executions.createComponents(Consts._CellMenupopup_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}

	public void onOpen() {
		DesktopSheetContext.getInstance(getDesktop()).reGainFocus();
	}
	
	public void onClick$cut() {
		DesktopSheetContext.getInstance(getDesktop()).cutSelection();
	}
	
	public void onClick$copy() {
		DesktopSheetContext.getInstance(getDesktop()).copySelection();
	}
	
	public void onClick$paste() {
		//EditHelper.doPaste(ss);
		DesktopSheetContext.getInstance(getDesktop()).pasteSelection();
	}
	
	public void onClick$pasteSpecial(MouseEvent event) {
		DesktopSheetContext.getInstance(getDesktop()).openPasteSpecialDialog();
	}

	public void onClick$shiftCellRight() {
		DesktopSheetContext.getInstance(getDesktop()).shiftCell(DesktopSheetContext.SHIFT_CELL_RIGHT);
	}
	
	public void onClick$shiftCellDown() {
		DesktopSheetContext.getInstance(getDesktop()).shiftCell(DesktopSheetContext.SHIFT_CELL_DOWN);
	}
	
	public void onClick$insertEntireRow() {
		DesktopSheetContext.getInstance(getDesktop()).insertRow();	
	}
	
	public void onClick$insertEntireColumn() {
		DesktopSheetContext.getInstance(getDesktop()).insertColumn();
	}
	
	public void onClick$shiftCellLeft() {
		DesktopSheetContext.getInstance(getDesktop()).shiftCell(DesktopSheetContext.SHIFT_CELL_LEFT);
	}
	
	public void onClick$shiftCellUp() {
		DesktopSheetContext.getInstance(getDesktop()).shiftCell(DesktopSheetContext.SHIFT_CELL_UP);
	}
	
	public void onClick$deleteEntireRow() {
		DesktopSheetContext.getInstance(getDesktop()).deleteRow();
	}

	public void onClick$deleteEntireColumn() {
		DesktopSheetContext.getInstance(getDesktop()).deleteColumn();
	}
	
	public void onClick$clearContent() {
		DesktopSheetContext.getInstance(getDesktop()).clearSelectionContent();
	}
	
	public void onClick$clearStyle() {
		DesktopSheetContext.getInstance(getDesktop()).clearSelectionStyle();
	}

	public void onClick$sortAscending() {
		DesktopSheetContext.getInstance(getDesktop()).sort(false);
	}
	
	public void onClick$sortDescending() {
		DesktopSheetContext.getInstance(getDesktop()).sort(true);
	}
	
	public void onClick$customSort() {
		DesktopSheetContext.getInstance(getDesktop()).openCustomSortDialog();
	}
	
	public void onClick$formula() {
		//open formula
		DesktopSheetContext.getInstance(getDesktop()).openInsertFormulaDialog();
	}

	public void onClick$format() {
		throw new UiException("not implement yet");
		//DesktopSheetContext.getInstance(getDesktop()).openFormatDialog();
	}
	
	public void onClick$hyperlink() {
		DesktopSheetContext.getInstance(getDesktop()).openHyperlinkDialog();
	}
}