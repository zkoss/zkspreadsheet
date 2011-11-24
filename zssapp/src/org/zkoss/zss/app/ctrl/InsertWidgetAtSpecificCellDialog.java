/* InsertWidgetAtCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2011 5:12:28 PM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.ctrl;

import java.util.ArrayList;
import java.util.HashMap;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.ext.AfterCompose;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Messagebox;

/**
 * @author sam
 * 
 */
public class InsertWidgetAtSpecificCellDialog extends Dialog implements AfterCompose {

	/*Views*/
	private Dialog insertWidgetAtDialog;
	private Combobox columns;
	private Combobox rows;
	private Button okBtn;
	
	/*Models*/
	private ListModelList availableColumns;
	private ListModelList availableRows;
	
	@Override
	public void afterCompose() {
		Components.wireVariables(this, this);
		Components.addForwards(this, this);
		
		buildHeaderModel();
		
		columns.setModel(availableColumns);
		rows.setModel(availableRows);
		
		final Rect visibleRect = getWorkbookCtrl().getVisibleRect();
		columns.addEventListener("onAfterRender", new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				int width = ((visibleRect.getRight() - visibleRect.getLeft() + 1) / 4);
				columns.setSelectedIndex(visibleRect.getLeft() + width);
			}
		});
		rows.addEventListener("onAfterRender", new EventListener() {

			@Override
			public void onEvent(Event event) throws Exception {
				int height = ((visibleRect.getBottom() - visibleRect.getTop() + 1) / 4);
				rows.setSelectedIndex(visibleRect.getTop() + height);
			}
		});
	}
	
	public void onOpen$insertWidgetAtDialog(ForwardEvent evt) {
		updateCenterColumnRow();
	}
	
	private void updateCenterColumnRow() {
		final Rect visibleRect = getWorkbookCtrl().getVisibleRect();
		int width = ((visibleRect.getRight() - visibleRect.getLeft() + 1) / 4);
		int col = visibleRect.getLeft() + width;
		if (col < columns.getChildren().size()) {
			columns.setSelectedIndex(visibleRect.getLeft() + width);
		}
		int height = ((visibleRect.getBottom() - visibleRect.getTop() + 1) / 4);
		int row = visibleRect.getTop() + height;
		if (row < rows.getChildren().size()) {
			rows.setSelectedIndex(visibleRect.getTop() + height);
		}
	}

	public void onClick$okBtn() {
		int col = columns.getSelectedIndex();
		int row = rows.getSelectedIndex();
		if (col < 0 || row < 0) {
			try {
				Messagebox.show("Please select row/column");
			} catch (InterruptedException e) {
			}
			return;
		}
		HashMap data = new HashMap();
		data.put("column", Integer.valueOf(col));
		data.put("row", Integer.valueOf(row));
		insertWidgetAtDialog.fireOnClose(data);
	}
	
	private void buildHeaderModel() {
		final WorkbookCtrl workbookCtrl = getWorkbookCtrl();
		
		int maxCol = getWorkbookCtrl().getMaxcolumns();
		ArrayList<String> strs = new ArrayList<String>();
		for (int i = 0; i < maxCol; i++) {
			strs.add(workbookCtrl.getColumnTitle(i));
		}
		availableColumns = new ListModelList(strs);
		
		int maxRow = getWorkbookCtrl().getMaxrows();
		strs = new ArrayList<String>();
		for (int i = 0; i < maxRow; i++) {
			strs.add(workbookCtrl.getRowTitle(i));
		}
		availableRows = new ListModelList(strs);
	}
	
	protected WorkbookCtrl getWorkbookCtrl() {
		return getDesktopWorkbenchContext().getWorkbookCtrl();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
}
