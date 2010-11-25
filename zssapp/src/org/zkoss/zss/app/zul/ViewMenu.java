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
	private Menuitem unfreezeRows;
	private Menupopup freezeColsMenupopup;
	private Menuitem unfreezeCols;
	
	public ViewMenu() {
		Executions.createComponents(Consts._ViewMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');

		getDesktopWorkbenchContext().addEventListener(Consts.ON_SHEET_OPEN, new EventListener() {
				public void onEvent(Event event) throws Exception {
					setDisabled(!(Boolean)event.getData());
				}
			});
	}
	
	public void onClick$viewFormulaBar() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().toggleFormulaBar();
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

	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop());
	}
}
