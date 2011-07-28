/* FormatMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 7:55:29 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import java.util.List;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.CellStyleContextEvent;
import org.zkoss.zss.app.zul.ctrl.DesktopCellStyleContext;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.StyleModification;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * @author Sam
 *
 */
public class FormatMenu extends Menu implements IdSpace {
	//TODO provide font color menu
	private Menupopup formatMenupopup;
	
	private Menupopup fontMenuMenupopup;
	private Menupopup alignMenupopup;
	private Menu backgroundColorMenu;
	private Menuitem formatNumber;

	public FormatMenu() {
		Executions.createComponents(Consts._FormatMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void onFontFamilySelect(ForwardEvent event) {
		final String fontFamily = (String)event.getData();
		getDesktopCellStyleContext().modifyStyle(new StyleModification() {
			public void modify(org.zkoss.zss.app.zul.ctrl.CellStyle style,
					CellStyleContextEvent candidteEvt) {
				style.setFontFamily(fontFamily);
			}
		});
	}
	
	public void setDisabled(boolean disable) {
		applyDisabled(fontMenuMenupopup.getChildren(), disable);
		//TODO: setContent shall disable, but it didn't
		backgroundColorMenu.setContent(disable ? "" : "#color=#FFFFFF");
		applyDisabled(alignMenupopup.getChildren(), disable);
		formatNumber.setDisabled(disable);
	}
	private void applyDisabled(List children, boolean disabled) {
		for (Object obj : children) {
			if (obj instanceof Menuitem) {
				Menuitem menu = (Menuitem)obj;
				menu.setDisabled(disabled);
			}
		}
	}
	
	public void onOpen$FormatMenu() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onChange$backgroundColorMenu(ForwardEvent event) {
		InputEvent evt = (InputEvent)event.getOrigin();
		final String color = (String)evt.getValue();
		getDesktopCellStyleContext().modifyStyle(new StyleModification() {
			public void modify(org.zkoss.zss.app.zul.ctrl.CellStyle style,
					CellStyleContextEvent candidteEvt) {
				style.setCellColor(color);
			}
		});
	}
	
	public void onAlignHorizontalClick(ForwardEvent event) {
		final String alignStr = (String) event.getData();

		getDesktopCellStyleContext().modifyStyle(new StyleModification() {
			public void modify(org.zkoss.zss.app.zul.ctrl.CellStyle style,
					CellStyleContextEvent candidteEvt) {
				short align = CellStyle.ALIGN_GENERAL;
				if (alignStr.equals("left")) {
					align = CellStyle.ALIGN_LEFT;
				} 

				if (alignStr.equals("center")) {
					align = CellStyle.ALIGN_CENTER;
				} 

				if (alignStr.equals("right")) {
					align = CellStyle.ALIGN_RIGHT;
				}
				style.setAlignment(align);
			}
		});
	}
	
	public void onOpen$formatMenupopup() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onClick$formatNumber() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openFormatNumberDialog();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
	
	protected DesktopCellStyleContext getDesktopCellStyleContext() {
		return Zssapp.getDesktopCellStyleContext(this);
	}

	public void onCreate() {
		final DesktopWorkbenchContext workbenchCtrl = getDesktopWorkbenchContext();
		workbenchCtrl.addEventListener(Consts.ON_WORKBOOK_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				setDisabled(!workbenchCtrl.getWorkbookCtrl().hasBook());
			}
		});
		if (WebApps.getFeature("pe"))
			backgroundColorMenu.setContent("#color=#FFFFFF");
		else {
			backgroundColorMenu.detach();
		}
	}
}