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
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopSheetContext;
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
		DesktopSheetContext.getInstance(getDesktop()).reGainFocus();
	}
	
	public RowHeaderMenupopup() {
		Executions.createComponents(Consts._RowHeaderMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void onClick$cut() {
		DesktopSheetContext.getInstance(getDesktop()).cutSelection();
	}
	
	public void onClick$copy() {
		DesktopSheetContext.getInstance(getDesktop()).copySelection();
	}
	
	public void onClick$paste() {
		DesktopSheetContext.getInstance(getDesktop()).pasteSelection();
	}
	
	public void onClick$clearContent() {
		DesktopSheetContext.getInstance(getDesktop()).clearSelectionContent();
	}
	
	public void onClick$clearStyle() {
		DesktopSheetContext.getInstance(getDesktop()).clearSelectionStyle();
	}
	
	public void onClick$insertRow() {
		DesktopSheetContext.getInstance(getDesktop()).insertRow();
	}
	
	public void onClick$deleteRow() {
		DesktopSheetContext.getInstance(getDesktop()).deleteRow();
	}
	
	public void onClick$rowHeight() {
		DesktopSheetContext.getInstance(getDesktop()).openModifyRowHeightDialog();
	}
	
	public void onClick$numberFormat() {
		throw new UiException("format not support yet");
	}
	
	public void onClick$hide() {
		DesktopSheetContext.getInstance(getDesktop()).hide(true);
	}
	
	public void onClick$unhide() {
		DesktopSheetContext.getInstance(getDesktop()).hide(false);
	}
}