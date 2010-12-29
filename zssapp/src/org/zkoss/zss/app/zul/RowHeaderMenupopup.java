/* RowHeaderMenupopup.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 24, 2010 11:16:23 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class RowHeaderMenupopup extends Menupopup implements IdSpace {

	private Menuitem cut;
	private Menuitem copy;
	private Menuitem paste;
	
	private Menuitem clearContent;
	private Menuitem clearStyle;
	
	private Menuitem insertRow;
	private Menuitem deleteRow;
	
	private Menuitem rowHeight;
	private Menuitem numberFormat;
	
	private Menuitem hide;
	private Menuitem unhide;
	
	public void onOpen() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public RowHeaderMenupopup() {
		Executions.createComponents(Consts._RowHeaderMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
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
	
	public void onClick$insertRow() {
		getDesktopWorkbenchContext().getWorkbookCtrl().insertRowAbove();
	}
	
	public void onClick$deleteRow() {
		getDesktopWorkbenchContext().getWorkbookCtrl().deleteRow();
	}
	
	public void onClick$rowHeight() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openModifyHeaderSizeDialog(WorkbookCtrl.HEADER_TYPE_ROW);
	}
	
	public void onClick$numberFormat() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openFormatNumberDialog();
	}
	
	public void onClick$hide() {
		getDesktopWorkbenchContext().getWorkbookCtrl().hide(true);
	}
	
	public void onClick$unhide() {
		getDesktopWorkbenchContext().getWorkbookCtrl().hide(false);
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
}