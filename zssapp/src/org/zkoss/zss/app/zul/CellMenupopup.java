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
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.CellStyle;
import org.zkoss.zss.app.zul.ctrl.CellStyleContext;
import org.zkoss.zss.app.zul.ctrl.CellStyleContextEvent;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.StyleModification;
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
	private Menuitem locked;
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
		System.out.println("rm onClick$customSort");
//		getDesktopWorkbookContext().getWorkbenchCtrl().openCustomSortDialog();
	}
	
	public void onClick$formula() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openInsertFormulaDialog();
	}

	public void onClick$format() {
		System.out.println("rm onClick$format");
//		getDesktopWorkbookContext().getWorkbenchCtrl().openFormatNumberDialog();
	}
	
	public void onCheck$locked() {
		
		getCellStyleContext().modifyStyle(new StyleModification(){
			public void modify(CellStyle style, CellStyleContextEvent candidteEvt) {
				candidteEvt.setExecutor(CellMenupopup.this);
				style.setLocked(locked.isChecked());
			}
		});
	}
	
	public void onClick$hyperlink() {
		System.out.println("rm onClick$hyperlink");
//		getDesktopWorkbookContext().getWorkbenchCtrl().openHyperlinkDialog();
	}
	
	private void init(CellStyle cellStyle) {
		locked.setDisabled(getDesktopWorkbenchContext().getWorkbookCtrl().isSheetProtect());
		locked.setChecked(cellStyle.getLocked());
	}
	
	public void onCreate() {
		CellStyleContext context = getCellStyleContext();
		DisposedEventListener listener = new DisposedEventListener() {
			public boolean isDisposed() {
				return CellMenupopup.this.getDesktop() == null;
			}
			public void onEvent(Event arg0) throws Exception {
				CellStyleContextEvent event = (CellStyleContextEvent) arg0;
				if(event.getExecutor() != CellMenupopup.this) {
					init(event.getCellStyle());
				}
			}
		};
		context.addEventListener(Consts.ON_STYLING_TARGET_CHANGED, listener);
		context.addEventListener(Consts.ON_CELL_STYLE_CHANGED, listener);
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbookContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
	
	protected CellStyleContext getCellStyleContext(){
		return Zssapp.getDesktopCellStyleContext(this);
	}
}