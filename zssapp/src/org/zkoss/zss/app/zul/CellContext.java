/* CellContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 7, 2010 10:31:11 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.CellStyleCtrlPanel;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class CellContext extends Window implements IdSpace {
	
	CellStyleCtrlPanel fontCtrlPanel;
	CellStyleCtrlPanel styleCtrlBottomPanel;
	Toolbarbutton _mergeCellBtn;

	public CellContext() {
		Executions.createComponents(Consts._CellContext_zul, this, null);
		
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
		setVisible(false);
		setSclass("fastIconWin");
		setVflex("min");
	}
	
	public void onClick$fontCtrlPanel() {
		setVisible(false);
	}
	
	public void onClick$styleCtrlBottomPanel() {
		setVisible(false);
	}
	
	public void onClick$_mergeCellBtn() {
		DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop()).mergeCell();
	}
}