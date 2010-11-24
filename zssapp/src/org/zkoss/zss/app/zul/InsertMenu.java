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
import org.zkoss.zss.app.zul.ctrl.DesktopSheetContext;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class InsertMenu extends Menu implements IdSpace {
	
	private final static String URI = "~./zssapp/html/menu/insertMenu.zul";
	
	private Menupopup insertMenupopup;
	
	private Menuitem insertFormula;
	private Menuitem insertChart;
	private Menuitem insertImage;
	private Menuitem insertSheet;

	
	public InsertMenu() {
		Executions.createComponents(URI, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	} 
	
	public void onClick$insertFormula() {
		DesktopSheetContext.getInstance(getDesktop()).openInsertFormulaDialog();
	}
	
	public void onClick$insertChart() {
		throw new UiException("insert Chart not implement yet");
	}
	
	public void onUpload$insertImage(ForwardEvent event) {
		insertMenupopup.close();
		UploadEvent evt = (UploadEvent)event.getOrigin();
		DesktopSheetContext.getInstance(getDesktop()).insertImage(evt.getMedia());
	}
	
	public void onClick$insertSheet() {
		DesktopSheetContext.getInstance(getDesktop()).insertSheet();
	}
	
	public void onOpen$insertMenupopup() {
		DesktopSheetContext.getInstance(getDesktop()).reGainFocus();
	}
}