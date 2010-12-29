/* DesktopSheetContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 2:51:14 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.app.Consts;

/**
 * @author Sam
 *
 */
public class DesktopWorkbenchContext extends AbstractBaseContext {

	public final static int SHIFT_CELL_UP = 0;
	public final static int SHIFT_CELL_RIGHT = 1;
	public final static int SHIFT_CELL_DOWN = 2;
	public final static int SHIFT_CELL_LEFT = 3;
	
	/* operate on selected sheet */
	WorkbookCtrl workbookCtrl;
	
	/**
	 * 
	 * @param workbookCtrl
	 */
	public void doTargetChange(WorkbookCtrl workbookCtrl) {
		this.workbookCtrl = workbookCtrl;
	}
	
	/**
	 * Returns the {@link #WorkbookCtrl}
	 * @return
	 */
	public WorkbookCtrl getWorkbookCtrl() {
		return workbookCtrl;
	}
	
	/* UI controls */
	WorkbenchCtrl workbenchCtrl;
	
	/**
	 * Sets the {@link #workbenchCtrl}
	 * @param workbenchCtrl
	 */
	public void setWorkbenchCtrl(WorkbenchCtrl workbenchCtrl) {
		this.workbenchCtrl = workbenchCtrl;
	}
	
	/**
	 * Returns the {@link #workbenchCtrl}
	 * @return
	 */
	public WorkbenchCtrl getWorkbenchCtrl() {
		return workbenchCtrl;
	}

	/**
	 * Fire sheet refresh event when sheet name changed
	 */
	public void fireRefresh() {
		listenerStore.fire(new Event(Consts.ON_SHEET_REFRESH));
	}

	/**
	 * Fire sheet changed event when spreadsheet set selected sheet
	 */
	public void fireSheetChanged() {
		listenerStore.fire(new Event(Consts.ON_SHEET_CHANGED));
	}
	public void fireContentsChanged() {
		listenerStore.fire(new Event(Consts.ON_SHEET_CONTENTS_CHANGED));
	}
	public void fireWorkbookSaved() {
		listenerStore.fire(new Event(Consts.ON_WORKBOOK_SAVED));
	}
	public void fireWorkbookChanged() {
		listenerStore.fire(new Event(Consts.ON_WORKBOOK_CHANGED));
	}
	public void mergeCell() {
		listenerStore.fire(new Event(Consts.ON_SHEET_MERGE_CELL));
	}
	public void insertFormula(String formula) {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT_FORMULA, null, formula));
	}
}