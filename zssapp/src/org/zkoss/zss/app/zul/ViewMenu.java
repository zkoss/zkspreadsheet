/* ViewMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 7:00:03 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import java.util.List;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class ViewMenu extends Menu implements IdSpace {

	private Menupopup viewMenupopup;
	
	private Menuitem viewFormulaBar;
	private Menupopup freezeRowsMenupopup;
	private Menupopup freezeColsMenupopup;
	
	private Menuitem unfreezeRows;
	private Menuitem unfreezeCols;
	
	private Menuitem freezeRow1;
	private Menuitem freezeRow2;
	private Menuitem freezeRow3;
	private Menuitem freezeRow4;
	private Menuitem freezeRow5;
	private Menuitem freezeRow6;
	private Menuitem freezeRow7;
	private Menuitem freezeRow8;
	private Menuitem freezeRow9;
	private Menuitem freezeRow10;
	
	private Menuitem freezeCol1;
	private Menuitem freezeCol2;
	private Menuitem freezeCol3;
	private Menuitem freezeCol4;
	private Menuitem freezeCol5;
	private Menuitem freezeCol6;
	private Menuitem freezeCol7;
	private Menuitem freezeCol8;
	private Menuitem freezeCol9;
	private Menuitem freezeCol10;
	
	public ViewMenu() {
		Executions.createComponents(Consts._ViewMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void setFreezeUnFreezeDisabled(boolean disabled) {
		unfreezeRows.setDisabled(disabled);
		unfreezeCols.setDisabled(disabled);
		
		freezeRow1.setDisabled(disabled);
		freezeRow2.setDisabled(disabled);
		freezeRow3.setDisabled(disabled);
		freezeRow4.setDisabled(disabled);
		freezeRow5.setDisabled(disabled);
		freezeRow6.setDisabled(disabled);
		freezeRow7.setDisabled(disabled);
		freezeRow8.setDisabled(disabled);
		freezeRow9.setDisabled(disabled);
		freezeRow10.setDisabled(disabled);
		
		freezeCol1.setDisabled(disabled);
		freezeCol2.setDisabled(disabled);
		freezeCol3.setDisabled(disabled);
		freezeCol4.setDisabled(disabled);
		freezeCol5.setDisabled(disabled);
		freezeCol6.setDisabled(disabled);
		freezeCol7.setDisabled(disabled);
		freezeCol8.setDisabled(disabled);
		freezeCol9.setDisabled(disabled);
		freezeCol10.setDisabled(disabled);
	}
	
	public void onClick$viewFormulaBar() {
		final boolean visible = getDesktopWorkbenchContext().getWorkbenchCtrl().toggleFormulaBar();
		viewFormulaBar.setChecked(visible);
	}
	
	public void onViewFreezeRows(ForwardEvent event) {
		getDesktopWorkbenchContext().getWorkbookCtrl().
			setRowFreeze(Integer.parseInt((String) event.getData()) - 1);
	}
	
	public void onViewFreezeCols(ForwardEvent event) {
		getDesktopWorkbenchContext().getWorkbookCtrl().
			setColumnFreeze(Integer.parseInt((String) event.getData()) - 1);
	}

	public void onClick$unfreezeRows(ForwardEvent event) {
		getDesktopWorkbenchContext().getWorkbookCtrl().setRowFreeze(-1);
	}

	public void onClick$unfreezeCols(ForwardEvent event) {
		getDesktopWorkbenchContext().getWorkbookCtrl().setColumnFreeze(-1);
	}
	
	public void setDisabled(boolean disabled) {
		viewFormulaBar.setDisabled(disabled);

		applyDisabled(freezeRowsMenupopup.getChildren(), disabled);
		applyDisabled(freezeColsMenupopup.getChildren(), disabled);
	}
	
	private void applyDisabled(List children, boolean disabled) {
		for (Object obj : children) {
			if (obj instanceof Menuitem) {
				Menuitem menu = (Menuitem)obj;
				menu.setDisabled(disabled);
			}
		}
	}
	
	public void onOpen$viewMenupopup() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}

	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}

	public void onCreate() {
		final DesktopWorkbenchContext workbenchCtrl = getDesktopWorkbenchContext();
		getDesktopWorkbenchContext().addEventListener(Consts.ON_WORKBOOK_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				setDisabled(!workbenchCtrl.getWorkbookCtrl().hasBook());
			}
		});
	}
}
