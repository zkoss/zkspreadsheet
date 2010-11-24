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
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopSheetContext;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class ViewMenu extends Menu implements ZssappComponent, IdSpace {

	private final static String URI = "~./zssapp/html/menu/viewMenu.zul";
	
	private Menupopup viewMenupopup;
	
	private Menuitem viewFormulaBar;
	
	private Menupopup freezeRowsMenupopup;
	private Menuitem unfreezeRows;
	
	private Menupopup freezeColsMenupopup;
	private Menuitem unfreezeCols;
	
	private Spreadsheet ss;
	
	public ViewMenu() {
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
	
	public void onClick$viewFormulaBar() {
		//TODO not implement yet
		throw new UiException("not implement yet");
	}
	
	public void onViewFreezeRows(ForwardEvent event) {
		ss.setRowfreeze(Integer.parseInt((String) event.getData()) - 1);
	}
	
	public void onViewFreezeCols(ForwardEvent event) {
		ss.setColumnfreeze(Integer.parseInt((String) event.getData()) - 1);
	}

	public void onClick$unfreezeRows(ForwardEvent event) {
		ss.setRowfreeze(-1);
	}

	public void onClick$unfreezeCols(ForwardEvent event) {
		ss.setColumnfreeze(-1);
	}
	
	public void setDisabled(boolean disabled) {
		viewFormulaBar.setDisabled(disabled);
		
//		for (Object obj : freezeRowsMenupopup.getChildren()) {
//			if (obj instanceof Menuitem) {
//				Menuitem menu = (Menuitem)obj;
//				menu.setDisabled(disabled);
//			}
//		}
		applyDisabled(freezeRowsMenupopup.getChildren(), disabled);
		
		applyDisabled(freezeColsMenupopup.getChildren(), disabled);
//		for (Object obj : freezeColsMenupopup.getChildren()) {
//			if (obj instanceof Menuitem) {
//				Menuitem menu = (Menuitem)obj;
//				menu.setDisabled(disabled);
//			}
//		}
	}
	
	private void applyDisabled(List children, boolean disabled) {
		for (Object obj : children) {
			if (obj instanceof Menuitem) {
				Menuitem menu = (Menuitem)obj;
				menu.setDisabled(disabled);
			}
		}
	}
	
	@Override
	public void bindSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
		viewMenupopup.setWidgetListener(Events.ON_OPEN, "this.$f('" + ss.getId() + "', true).focus(false);");
	}

	@Override
	public void unbindSpreadsheet() {
		// TODO Auto-generated method stub
		
	}

}
