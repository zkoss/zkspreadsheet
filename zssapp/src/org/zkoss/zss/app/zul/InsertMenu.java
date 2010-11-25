/* InsertMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 8:38:14 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class InsertMenu extends Menu implements IdSpace {

	
	private Menupopup insertMenupopup;	
	private Menuitem insertFormula;
	private Menuitem insertChart;
	private Menuitem insertImage;
	private Menuitem insertSheet;

	
	public InsertMenu() {
		Executions.createComponents(Consts._InsertMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	} 
	
	public void onClick$insertFormula() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openInsertFormulaDialog();
	}
	
	public void onClick$insertChart() {
		throw new UiException("insert Chart not implement yet");
	}
	
	public void onUpload$insertImage(ForwardEvent event) {
		insertMenupopup.close();
		UploadEvent evt = (UploadEvent)event.getOrigin();
		getDesktopWorkbenchContext().getWorkbookCtrl().insertImage(evt.getMedia());
	}
	
	public void onClick$insertSheet() {
		getDesktopWorkbenchContext().getWorkbookCtrl().insertSheet();
		getDesktopWorkbenchContext().fireRefresh();
	}
	
	public void onOpen$insertMenupopup() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop());
	}
}