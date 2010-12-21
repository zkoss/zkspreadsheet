/* ColumnHeaderMenupopup.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 24, 2010 2:02:56 PM , Created by Sam
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
public class ColumnHeaderMenupopup  extends Menupopup implements IdSpace {
	private Menuitem cut;
	private Menuitem copy;
	private Menuitem paste;
	
	private Menuitem clearContent;
	private Menuitem clearStyle;
	
	private Menuitem insertColumn;
	private Menuitem deleteColumn;
	
	private Menuitem columnWidth;
	private Menuitem numberFormat;
	
	private Menuitem hide;
	private Menuitem unhide;
	
	public void onOpen() {
		getDesktopWorkbookContext().getWorkbookCtrl().reGainFocus();
	}
	
	public ColumnHeaderMenupopup() {
		Executions.createComponents(Consts._ColumnHeaderMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
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
	
	public void onClick$clearContent() {
		getDesktopWorkbookContext().getWorkbookCtrl().clearSelectionContent();
	}
	
	public void onClick$clearStyle() {
		getDesktopWorkbookContext().getWorkbookCtrl().clearSelectionStyle();
	}
	
	public void onClick$insertColumn() {
		getDesktopWorkbookContext().getWorkbookCtrl().insertColumnLeft();
	}
	
	public void onClick$deleteColumn() {
		getDesktopWorkbookContext().getWorkbookCtrl().deleteColumn();
	}
	
	public void onClick$columnWidth() {
		DesktopWorkbenchContext.getInstance(
			Executions.getCurrent().getDesktop()).getWorkbenchCtrl().openModifyHeaderSizeDialog(WorkbookCtrl.HEADER_TYPE_COLUMN);
	}
	
	public void onClick$numberFormat() {
		getDesktopWorkbookContext().getWorkbenchCtrl().openFormatNumberDialog();
	}
	
	public void onClick$hide() {
		getDesktopWorkbookContext().getWorkbookCtrl().hide(true);
	}
	
	public void onClick$unhide() {
		getDesktopWorkbookContext().getWorkbookCtrl().hide(false);
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbookContext() {
		return DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop());
	}
}