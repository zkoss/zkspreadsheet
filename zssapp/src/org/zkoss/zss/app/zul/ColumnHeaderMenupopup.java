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
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopSheetContext;
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
		DesktopSheetContext.getInstance(getDesktop()).reGainFocus();
	}
	
	public ColumnHeaderMenupopup() {
		Executions.createComponents(Consts._ColumnHeaderMenu_zul, this, null);
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
	
	public void onClick$insertColumn() {
		DesktopSheetContext.getInstance(getDesktop()).insertColumn();
	}
	
	public void onClick$deleteColumn() {
		DesktopSheetContext.getInstance(getDesktop()).deleteColumn();
	}
	
	public void onClick$columnWidth() {
		//TODO
		throw new UiException("not implement yet");
	}
	
	public void onClick$numberFormat() {
		//TODO
		throw new UiException("not implement yet");
	}
	
	public void onClick$hide() {
		DesktopSheetContext.getInstance(getDesktop()).hide(true);
	}
	
	public void onClick$unhide() {
		DesktopSheetContext.getInstance(getDesktop()).hide(false);
	}
}